; Instalador do Painel da Granja — Inno Setup
[Setup]
AppName=Painel da Granja
AppVersion=1.0
AppPublisher=Painel da Granja
DefaultDirName={localappdata}\Programs\Painel da Granja
DefaultGroupName=Painel da Granja
UninstallDisplayName=Painel da Granja
UninstallDisplayIcon={app}\PainelGranja.exe
SetupIconFile=D:\guia-ia\painel-granja\icone.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
OutputDir=D:\guia-ia\painel-granja\instalador
OutputBaseFilename=Instalador - Painel da Granja
LicenseFile=

[Languages]
Name: "ptbr"; MessagesFile: "C:\Users\orolg\AppData\Local\Programs\Inno Setup 6\Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho no Desktop"; \
    GroupDescription: "Atalhos:"; Flags: unchecked
Name: "iniciowindows"; Description: "Abrir o Painel automaticamente toda vez que ligar o computador (recomendado)"; \
    GroupDescription: "Início automático:"; Flags: checkedonce

[Files]
Source: "D:\guia-ia\painel-granja\dist\PainelGranja.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\guia-ia\painel-granja\LEIA-ME.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Painel da Granja"; Filename: "{app}\PainelGranja.exe"
Name: "{group}\Desinstalar Painel da Granja"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Painel da Granja"; Filename: "{app}\PainelGranja.exe"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; \
    ValueName: "Painel da Granja"; ValueData: """{app}\PainelGranja.exe"""; Tasks: iniciowindows; \
    Flags: uninsdeletevalue

[Run]
Filename: "{app}\PainelGranja.exe"; Description: "Abrir o Painel da Granja agora"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; o desinstalador NÃO apaga os dados do usuário (ficam em %APPDATA%\PainelGranja)
