; Inno Setup Script — 名簿帳票ツール
;
; ビルド:
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=v1.0.3 installer\setup.iss
;
; 出力:
;   dist\MeiboToolSetup-{AppVersion}.exe

#ifndef AppVersion
  #define AppVersion "v0.0.0"
#endif

[Setup]
AppId={{B7F3E2A1-4D8C-4F9B-A6E0-1C2D3E4F5A6B}
AppName=名簿帳票ツール
AppVersion={#AppVersion}
AppVerName=名簿帳票ツール {#AppVersion}
AppPublisher=名簿帳票ツール
DefaultDirName={userdocs}\名簿帳票ツール
DisableProgramGroupPage=yes
DisableDirPage=no
OutputDir=..\dist
OutputBaseFilename=MeiboToolSetup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
UninstallDisplayName=名簿帳票ツール
PrivilegesRequired=lowest
SetupIconFile=
WizardStyle=modern

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Files]
; メイン exe
Source: "..\dist\名簿帳票ツール\名簿帳票ツール.exe"; DestDir: "{app}"; Flags: ignoreversion

; _internal ディレクトリ（バンドルリソース全体）
Source: "..\dist\名簿帳票ツール\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

; config.json — ユーザー編集版がなければ初期版をコピー
Source: "..\dist\名簿帳票ツール\_internal\config.json"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
; デスクトップショートカット
Name: "{userdesktop}\名簿帳票ツール"; Filename: "{app}\名簿帳票ツール.exe"; WorkingDir: "{app}"

[Run]
; インストール完了後にアプリを起動するオプション
Filename: "{app}\名簿帳票ツール.exe"; Description: "名簿帳票ツールを起動"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; アンインストール時に _internal を削除（ユーザーデータは残す）
Type: filesandordirs; Name: "{app}\_internal"
