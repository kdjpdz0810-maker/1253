# -*- coding: utf-8 -*-
"""
그림파일 변환기
 - 문서(hwpx, docx, xlsx, pptx, pdf, html) -> 그림(png, jpg)
 - 첫 페이지 / 모든 페이지 선택, 저장 위치 = 원본 문서가 있는 폴더
 - 모든 페이지 저장 시 파일명에 자동 페이지 번호(_001, _002 ...)
 - 배경 제거(투명 PNG) 기능 (문서 변환 시 옵션 / 이미지 파일 단독 처리)

변환 방식 (Windows 기준)
 - PDF   : PyMuPDF 로 직접 렌더링
 - DOCX/XLSX/PPTX : MS Office(COM) -> PDF -> 그림   (실패 시 LibreOffice)
 - HWPX  : 한컴오피스(COM) -> PDF -> 그림           (실패 시 텍스트 전용 대체 변환)
 - HTML  : Microsoft Edge / Chrome (headless) -> PDF -> 그림
"""
import os
import re
import sys
import queue
import shutil
import zipfile
import tempfile
import threading
import subprocess
import traceback
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
from PIL import Image
import fitz  # PyMuPDF

try:
    from scipy import ndimage
except Exception:  # scipy 없으면 '전체 같은 색' 방식만 사용
    ndimage = None

DOC_EXTS = {".hwpx", ".docx", ".xlsx", ".pptx", ".pdf", ".html", ".htm"}
IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
ALL_EXTS = DOC_EXTS | IMG_EXTS
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


# ----------------------------------------------------------------------
# 1) 문서 -> PDF 변환
# ----------------------------------------------------------------------
def office_to_pdf(src: Path, dst: Path):
    """MS Office COM 자동화 (Word / Excel / PowerPoint)"""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    ext = src.suffix.lower()
    s, d = os.path.abspath(src), os.path.abspath(dst)
    try:
        if ext == ".docx":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                doc = app.Documents.Open(s, ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(d, 17)  # wdExportFormatPDF
                doc.Close(False)
            finally:
                app.Quit()
        elif ext == ".xlsx":
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(s, ReadOnly=True)
                wb.ExportAsFixedFormat(0, d)  # xlTypePDF
                wb.Close(False)
            finally:
                app.Quit()
        elif ext == ".pptx":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                pres = app.Presentations.Open(s, ReadOnly=-1, WithWindow=0)
                pres.SaveAs(d, 32)  # ppSaveAsPDF
                pres.Close()
            finally:
                app.Quit()
    finally:
        pythoncom.CoUninitialize()
    if not dst.exists():
        raise RuntimeError("Office 에서 PDF 파일이 생성되지 않았습니다.")


def find_libreoffice():
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(c):
            return c
    return None


def libreoffice_to_pdf(src: Path, dst: Path):
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("MS Office 와 LibreOffice 를 모두 찾을 수 없습니다.")
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, str(src)],
                       check=True, timeout=180, creationflags=NO_WINDOW,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        out = Path(td) / (src.stem + ".pdf")
        if not out.exists():
            raise RuntimeError("LibreOffice 변환 실패")
        shutil.move(str(out), str(dst))


def find_browser():
    cands = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    return shutil.which("msedge") or shutil.which("chrome") or shutil.which("google-chrome") \
        or shutil.which("chromium")


def html_to_pdf(src: Path, dst: Path):
    browser = find_browser()
    if not browser:
        raise RuntimeError("Microsoft Edge 또는 Chrome 을 찾을 수 없습니다.")
    with tempfile.TemporaryDirectory() as profile:
        cmd = [browser, "--headless", "--disable-gpu", "--no-sandbox",
               "--no-pdf-header-footer", "--print-to-pdf-no-header",
               f"--user-data-dir={profile}",
               f"--print-to-pdf={os.path.abspath(dst)}",
               src.resolve().as_uri()]
        try:
            subprocess.run(cmd, timeout=90, creationflags=NO_WINDOW,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except subprocess.TimeoutExpired:
            pass
    if not dst.exists():
        raise RuntimeError("HTML -> PDF 변환 실패")


# ---- HWPX ------------------------------------------------------------
def find_hwp_security_dll():
    """한컴 '파일 경로 접근 허용' 보안모듈 DLL 찾기 (exe 와 같은 폴더 또는 hwp_security 폴더)"""
    for base in (app_dir(), app_dir() / "hwp_security"):
        if base.exists():
            for f in base.glob("FilePathChecker*.dll"):
                return f
    return None


def register_hwp_security(hwp):
    """한컴 자동화 접근 권한 허용 (보안승인 팝업 방지)"""
    dll = find_hwp_security_dll()
    if dll:
        try:
            import winreg
            for key in (r"Software\HNC\HwpAutomation\Modules", r"Software\HNC\HwpCtrl\Modules"):
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as k:
                    winreg.SetValueEx(k, "FilePathCheckerModule", 0, winreg.REG_SZ, str(dll))
        except Exception:
            pass
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass


def hwpx_to_pdf(src: Path, dst: Path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        register_hwp_security(hwp)
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        if not hwp.Open(os.path.abspath(src), "HWPX", "forceopen:true"):
            raise RuntimeError("한글에서 파일을 열 수 없습니다 (접근 권한 확인).")
        if not hwp.SaveAs(os.path.abspath(dst), "PDF", ""):
            raise RuntimeError("한글에서 PDF 저장에 실패했습니다.")
    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    if not dst.exists():
        raise RuntimeError("PDF 파일이 생성되지 않았습니다.")


def _hwpx_paragraphs(src: Path):
    paras = []
    with zipfile.ZipFile(src) as z:
        sections = sorted((n for n in z.namelist() if re.match(r"Contents/section\d+\.xml$", n)),
                          key=lambda n: int(re.findall(r"\d+", n)[-1]))
        for name in sections:
            root = ET.fromstring(z.read(name))
            out = []

            def walk(el, buf):
                tag = el.tag.split("}")[-1]
                if tag == "p":
                    buf = []
                    out.append(buf)
                elif tag == "t" and buf is not None:
                    buf.append("".join(el.itertext()))
                for ch in el:
                    walk(ch, buf)

            walk(root, None)
            paras += ["".join(b) for b in out if "".join(b).strip()]
    return paras


def hwpx_text_fallback(src: Path, dst: Path):
    """한글 프로그램이 없을 때: 본문 텍스트만 추출하여 PDF 생성 (서식/이미지 제외)"""
    paras = _hwpx_paragraphs(src)
    if not paras:
        raise RuntimeError("HWPX 에서 텍스트를 읽지 못했습니다.")
    fontfile = None
    for f in (r"C:\Windows\Fonts\malgun.ttf", r"C:\Windows\Fonts\gulim.ttc",
              "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"):
        if os.path.exists(f):
            fontfile = f
            break
    if not fontfile:
        raise RuntimeError("한글 글꼴을 찾을 수 없습니다.")
    font = fitz.Font(fontfile=fontfile)
    W, H, M, size, lh = 595, 842, 56, 11, 17
    doc = fitz.open()
    page = doc.new_page(width=W, height=H)
    y = M
    for para in paras:
        line, lines = "", []
        for ch in para:
            if font.text_length(line + ch, fontsize=size) > W - 2 * M:
                lines.append(line)
                line = ch
            else:
                line += ch
        lines.append(line)
        for ln in lines:
            if y > H - M:
                page = doc.new_page(width=W, height=H)
                y = M
            page.insert_text((M, y), ln, fontsize=size, fontname="kfont", fontfile=fontfile)
            y += lh
        y += 4
    doc.save(str(dst))
    doc.close()


def to_pdf(src: Path, dst: Path, log):
    ext = src.suffix.lower()
    if ext == ".hwpx":
        try:
            hwpx_to_pdf(src, dst)
        except Exception as e:
            log(f"  ! 한글 변환 실패: {e}")
            log("  → 텍스트 전용 대체 변환 사용 (서식·이미지 제외)")
            hwpx_text_fallback(src, dst)
    elif ext in (".html", ".htm"):
        html_to_pdf(src, dst)
    elif ext in (".docx", ".xlsx", ".pptx"):
        try:
            office_to_pdf(src, dst)
        except Exception as e:
            log(f"  ! MS Office 변환 실패: {e}")
            log("  → LibreOffice 로 재시도")
            libreoffice_to_pdf(src, dst)
    else:
        raise RuntimeError(f"지원하지 않는 형식: {ext}")


# ----------------------------------------------------------------------
# 2) 배경 제거 (투명 PNG)
# ----------------------------------------------------------------------
def remove_background(img: Image.Image, tol: int = 30, edge_only: bool = True) -> Image.Image:
    """
    테두리 픽셀의 대표색(문서는 보통 흰색)을 배경색으로 보고 투명 처리.
    edge_only=True : 가장자리와 연결된 배경만 제거 (글자 안쪽 흰색 등은 유지)
    edge_only=False: 이미지 전체에서 배경색과 비슷한 픽셀 모두 제거
    """
    arr = np.array(img.convert("RGBA"))
    rgb = arr[:, :, :3].astype(np.int16)
    border = np.concatenate([rgb[0], rgb[-1], rgb[:, 0], rgb[:, -1]])
    bg = np.median(border, axis=0)
    cand = np.abs(rgb - bg).max(axis=2) <= tol

    if edge_only and ndimage is not None:
        labels, _ = ndimage.label(cand)  # 4방향 연결
        edge_labels = np.unique(np.concatenate(
            [labels[0], labels[-1], labels[:, 0], labels[:, -1]]))
        edge_labels = edge_labels[edge_labels != 0]
        mask = np.isin(labels, edge_labels)
    else:
        mask = cand

    arr[:, :, 3] = np.where(mask, 0, arr[:, :, 3])
    return Image.fromarray(arr, "RGBA")


# ----------------------------------------------------------------------
# 3) PDF -> 그림
# ----------------------------------------------------------------------
def pdf_to_images(pdf: Path, out_dir: Path, stem: str, opts: dict):
    fmt = opts["fmt"]  # png / jpg
    transparent = opts["transparent"]
    if transparent:
        fmt = "png"  # JPG 는 투명 미지원
    ext = ".png" if fmt == "png" else ".jpg"
    zoom = opts["dpi"] / 72.0

    doc = fitz.open(str(pdf))
    if doc.needs_pass:
        raise RuntimeError("암호가 걸린 PDF 입니다.")
    total = doc.page_count
    count = 1 if opts["mode"] == "first" else total
    pad = max(3, len(str(total)))
    outputs = []
    for i in range(count):
        pix = doc[i].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        if transparent:
            img = remove_background(img, opts["tol"], opts["edge_only"])
        name = f"{stem}{ext}" if opts["mode"] == "first" else f"{stem}_{i + 1:0{pad}d}{ext}"
        out = out_dir / name
        if fmt == "png":
            img.save(out, "PNG")
        else:
            img.convert("RGB").save(out, "JPEG", quality=95)
        outputs.append(out)
    doc.close()
    return outputs


def process_file(path: Path, opts: dict, log):
    out_dir = path.parent  # 저장 위치 = 문서가 있는 폴더
    ext = path.suffix.lower()
    if ext in IMG_EXTS:  # 이미지 파일 -> 배경 제거 후 투명 PNG
        img = Image.open(path)
        res = remove_background(img, opts["tol"], opts["edge_only"])
        out = out_dir / f"{path.stem}_투명.png"
        res.save(out, "PNG")
        return [out]

    with tempfile.TemporaryDirectory() as td:
        if ext == ".pdf":
            pdf = path
        else:
            pdf = Path(td) / "converted.pdf"
            to_pdf(path, pdf, log)
        return pdf_to_images(pdf, out_dir, path.stem, opts)


# ----------------------------------------------------------------------
# 4) GUI
# ----------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        has_dnd = True
    except Exception:
        has_dnd = False

    root = TkinterDnD.Tk() if has_dnd else tk.Tk()
    root.title("그림파일 변환기")
    root.geometry("860x800")
    root.minsize(760, 700)

    files, q = [], queue.Queue()
    fmt_var = tk.StringVar(value="png")
    mode_var = tk.StringVar(value="first")
    dpi_var = tk.StringVar(value="200")
    trans_var = tk.BooleanVar(value=False)
    bg_var = tk.StringVar(value="가장자리와 연결된 배경만 제거")
    tol_var = tk.IntVar(value=30)

    pad = {"padx": 12, "pady": 6}

    # --- 업로드 박스
    drop = tk.Label(
        root, bg="#eef3fb", bd=2, relief="groove", height=6, cursor="hand2",
        font=("맑은 고딕", 11),
        text=("📂  여기로 파일을 끌어다 놓거나, 클릭하여 파일을 선택하세요\n\n"
              "문서: hwpx · docx · xlsx · pptx · pdf · html\n"
              "이미지(png·jpg): 배경 제거(투명 PNG) 전용"
              + ("" if has_dnd else "\n※ 드래그 앤 드롭 모듈 없음 - 클릭 선택만 가능")))
    drop.pack(fill="x", **pad)

    # --- 파일 목록
    frm = ttk.Frame(root)
    frm.pack(fill="both", expand=True, **pad)
    tree = ttk.Treeview(frm, columns=("name", "type", "status"), show="headings", height=7,
                        selectmode="extended")
    for c, t, w in (("name", "파일 (저장 위치 = 같은 폴더)", 470), ("type", "형식", 70),
                    ("status", "상태", 200)):
        tree.heading(c, text=t)
        tree.column(c, width=w, anchor="w")
    sb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    def add_files(paths):
        for p in paths:
            p = Path(p)
            if p.is_dir():
                continue
            if p.suffix.lower() not in ALL_EXTS:
                log(f"지원하지 않는 형식 제외: {p.name}")
                continue
            if p in files:
                continue
            files.append(p)
            tree.insert("", "end", iid=str(p),
                        values=(str(p), p.suffix.lower().lstrip("."), "대기"))

    def pick():
        exts = " ".join(f"*{e}" for e in sorted(ALL_EXTS))
        sel = filedialog.askopenfilenames(
            title="파일 선택",
            filetypes=[("지원 파일", exts), ("문서", "*.hwpx *.docx *.xlsx *.pptx *.pdf *.html *.htm"),
                       ("이미지", "*.png *.jpg *.jpeg *.bmp *.webp"), ("모든 파일", "*.*")])
        add_files(sel)

    def on_drop(e):
        add_files(root.tk.splitlist(e.data))

    drop.bind("<Button-1>", lambda e: pick())
    if has_dnd:
        for w in (drop, tree):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", on_drop)

    bar = ttk.Frame(root)
    bar.pack(fill="x", padx=12)

    def remove_sel():
        for iid in tree.selection():
            files.remove(Path(iid))
            tree.delete(iid)

    def clear_all():
        files.clear()
        tree.delete(*tree.get_children())

    def open_folder():
        sel = tree.selection()
        target = Path(sel[0]).parent if sel else (files[0].parent if files else None)
        if target and hasattr(os, "startfile"):
            os.startfile(target)

    ttk.Button(bar, text="파일 추가", command=pick).pack(side="left")
    ttk.Button(bar, text="선택 삭제", command=remove_sel).pack(side="left", padx=4)
    ttk.Button(bar, text="목록 비우기", command=clear_all).pack(side="left")
    ttk.Button(bar, text="저장 폴더 열기", command=open_folder).pack(side="right")

    # --- 옵션
    opt = ttk.LabelFrame(root, text=" 변환 옵션 ")
    opt.pack(fill="x", **pad)

    r1 = ttk.Frame(opt)
    r1.pack(fill="x", padx=8, pady=4)
    ttk.Label(r1, text="그림 형식 :", width=11).pack(side="left")
    ttk.Radiobutton(r1, text="PNG", value="png", variable=fmt_var).pack(side="left", padx=4)
    rb_jpg = ttk.Radiobutton(r1, text="JPG", value="jpg", variable=fmt_var)
    rb_jpg.pack(side="left", padx=4)
    ttk.Label(r1, text="해상도(DPI) :").pack(side="left", padx=(30, 4))
    ttk.Combobox(r1, textvariable=dpi_var, values=("96", "150", "200", "300"),
                 width=6, state="readonly").pack(side="left")

    r2 = ttk.Frame(opt)
    r2.pack(fill="x", padx=8, pady=4)
    ttk.Label(r2, text="저장 방법 :", width=11).pack(side="left")
    ttk.Radiobutton(r2, text="첫 페이지만", value="first", variable=mode_var).pack(side="left", padx=4)
    ttk.Radiobutton(r2, text="모든 페이지 (파일명_001, _002 …)", value="all",
                    variable=mode_var).pack(side="left", padx=4)

    r3 = ttk.Frame(opt)
    r3.pack(fill="x", padx=8, pady=4)
    ttk.Checkbutton(r3, text="배경을 지우고 투명 PNG 로 저장", variable=trans_var).pack(side="left")
    ttk.Combobox(r3, textvariable=bg_var, width=30, state="readonly",
                 values=("가장자리와 연결된 배경만 제거", "배경색과 같은 색 전체 제거")
                 ).pack(side="left", padx=10)

    r4 = ttk.Frame(opt)
    r4.pack(fill="x", padx=8, pady=(0, 6))
    ttk.Label(r4, text="배경 허용오차 :", width=14).pack(side="left")
    tol_lbl = ttk.Label(r4, text="30", width=4)
    ttk.Scale(r4, from_=0, to=120, variable=tol_var, length=260,
              command=lambda v: tol_lbl.config(text=str(int(float(v))))).pack(side="left")
    tol_lbl.pack(side="left", padx=6)
    ttk.Label(r4, text="(클수록 더 많이 지워짐 / 글자 주변 흰 테두리가 남으면 값을 올리세요)",
              foreground="#666").pack(side="left")

    def on_trans(*_):
        if trans_var.get():
            fmt_var.set("png")
            rb_jpg.state(["disabled"])
        else:
            rb_jpg.state(["!disabled"])

    trans_var.trace_add("write", on_trans)

    # --- 실행/진행/로그
    go_btn = ttk.Button(root, text="▶ 변환 시작")
    go_btn.pack(fill="x", padx=12, pady=(4, 2), ipady=6)
    prog = ttk.Progressbar(root, mode="determinate")
    prog.pack(fill="x", padx=12, pady=2)
    logbox = tk.Text(root, height=8, state="disabled", bg="#fafafa", font=("Consolas", 9))
    logbox.pack(fill="both", expand=False, padx=12, pady=(2, 10))

    def log(msg):
        logbox.config(state="normal")
        logbox.insert("end", msg + "\n")
        logbox.see("end")
        logbox.config(state="disabled")

    def worker(items, o):
        for i, p in enumerate(items, 1):
            q.put(("status", p, "변환 중..."))
            try:
                outs = process_file(p, o, lambda m: q.put(("log", m)))
                q.put(("log", f"[완료] {p.name} → {len(outs)}개 저장 ({p.parent})"))
                q.put(("status", p, f"완료 ({len(outs)}개)"))
            except Exception as e:
                q.put(("log", f"[오류] {p.name}: {e}"))
                q.put(("log", traceback.format_exc()))
                q.put(("status", p, f"오류: {e}"[:60]))
            q.put(("prog", i, len(items)))
        q.put(("done",))

    def start():
        if not files:
            messagebox.showinfo("안내", "변환할 파일을 먼저 추가하세요.")
            return
        o = {"fmt": fmt_var.get(), "mode": mode_var.get(), "dpi": int(dpi_var.get()),
             "transparent": trans_var.get(), "tol": int(tol_var.get()),
             "edge_only": bg_var.get().startswith("가장자리")}
        if o["transparent"]:
            log("※ 투명 배경은 PNG 로만 저장됩니다.")
        go_btn.state(["disabled"])
        prog.config(value=0, maximum=len(files))
        threading.Thread(target=worker, args=(list(files), o), daemon=True).start()

    go_btn.config(command=start)

    def poll():
        try:
            while True:
                m = q.get_nowait()
                if m[0] == "log":
                    log(m[1])
                elif m[0] == "status":
                    if tree.exists(str(m[1])):
                        tree.set(str(m[1]), "status", m[2])
                elif m[0] == "prog":
                    prog.config(value=m[1])
                elif m[0] == "done":
                    go_btn.state(["!disabled"])
                    log("--- 모든 작업이 끝났습니다 ---")
        except queue.Empty:
            pass
        root.after(100, poll)

    poll()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
