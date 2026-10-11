\xEF\xBB\xBF[Setup]
AppName=그림파일 변환기
AppVersion=1.0
AppPublisher=PictureConverter
DefaultDirName={autopf}\PictureConverter
DefaultGroupName=그림파일 변환기
OutputDir=installer_output
OutputBaseFilename=PictureConverter_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\PictureConverter.exe

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; GroupDescription: "추가 작업:"

[Files]
Source: "dist\PictureConverter.exe"; DestDir: "{app}"; Flags: ignoreversion
; 한컴 보안모듈 DLL 을 hwp_security 폴더에 넣어 두면 함께 설치됩니다
Source: "hwp_security\*"; DestDir: "{app}\hwp_security"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\그림파일 변환기"; Filename: "{app}\PictureConverter.exe"
Name: "{autodesktop}\그림파일 변환기"; Filename: "{app}\PictureConverter.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PictureConverter.exe"; Description: "지금 실행"; Flags: nowait postinstall skipifsilent
