#define MyAppName "Gerador de Diário de Obras"
#define MyAppVersion "1.2.1"
#define MyAppPublisher "Assessoria Tech"
#define MyAppExeName "GeradorDiarioObra.exe"
#define MyAppFolder "GeradorDiarioObra"

[Setup]
AppId={{D7C0DBA5-6C6A-4A91-8D74-9A9D4E9F1101}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppFolder}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=no
AllowNoIcons=yes
OutputDir=instalador
OutputBaseFilename=Setup_GeradorDiarioObra_v1_2_1
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=icone.ico
UninstallDisplayIcon={app}\icone.ico
UsePreviousAppDir=yes
CloseApplications=no

[Languages]
Name: "portuguesebrazil"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Opções adicionais:"; Flags: unchecked

[Dirs]
Name: "{app}"
Name: "{app}\templates"
Name: "{localappdata}\GeradorDiarioObra"
Name: "{localappdata}\GeradorDiarioObra\logs"

[Files]
Source: "dist\GeradorDiarioObra.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "templates\modelopadrao.docx"; DestDir: "{app}\templates"; Flags: ignoreversion
Source: "icone.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\icone.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\icone.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{localappdata}\GeradorDiarioObra\logs\erro.log"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if not DirExists(ExpandConstant('{localappdata}\GeradorDiarioObra')) then
      CreateDir(ExpandConstant('{localappdata}\GeradorDiarioObra'));

    if not DirExists(ExpandConstant('{localappdata}\GeradorDiarioObra\logs')) then
      CreateDir(ExpandConstant('{localappdata}\GeradorDiarioObra\logs'));
  end;
end;