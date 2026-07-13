; ================================================================
;  FERRAMENTA EACE NOC — Instalador Oficial Bitnet Telecom
;  Inno Setup 6 Script
;  Para atualizar: incremente MyAppVersion e recompile
; ================================================================

#define MyAppName      "Ferramenta EACE NOC"
#define MyAppVersion   "8.6"
#define MyAppPublisher "Bitnet Telecom"
#define MyAppExeName   "Ferramenta_EACE_NOC.exe"
#define SourceDir      "."

[Setup]
; ATENÇÃO: Mantenha o mesmo AppId em todas as versões para que o Windows
; reconheça atualizações como o mesmo aplicativo (substitui, não duplica).
AppId={{F8A4D2B1-3C7E-4F9A-B015-6D2E8C1A9F34}

AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} v{#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://bitnet.com.br
AppSupportURL=https://bitnet.com.br
AppUpdatesURL=https://bitnet.com.br

; Instala na pasta do usuário — não exige administrador
DefaultDirName={localappdata}\Programs\FeramentaEACEBitnet
DefaultGroupName={#MyAppName}
AllowNoIcons=yes

; Instalação por usuário (sem UAC)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Saída do instalador compilado
OutputDir={#SourceDir}\installer_output
OutputBaseFilename=Ferramenta_EACE_NOC_Setup_v{#MyAppVersion}

; Ícone do instalador (gerado pelo build_installer.ps1)
SetupIconFile={#SourceDir}\logo_bitnet.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Compressão máxima
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern

; Fecha o aplicativo automaticamente ao atualizar
CloseApplications=yes
CloseApplicationsFilter=*{#MyAppExeName}*
RestartApplications=no

; Metadados da versão do instalador
VersionInfoVersion=5.0.0.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} - Ferramenta de Automação NOC
VersionInfoCopyright=© 2025 Bitnet Telecom
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos adicionais:"; Flags: unchecked

[Files]
; ---------------------------------------------------------------
; EXECUTÁVEL PRINCIPAL
; Sempre substituído em atualizações (versão mais recente)
; ---------------------------------------------------------------
Source: "{#SourceDir}\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; ---------------------------------------------------------------
; TEMPLATES DE CONTRATO
; Sempre substituídos (garantir versão mais recente dos modelos)
; ---------------------------------------------------------------
Source: "{#SourceDir}\contrato modeloBIT.docx"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\contrato modeloST1.docx"; DestDir: "{app}"; Flags: ignoreversion

; ---------------------------------------------------------------
; CREDENCIAIS GOOGLE SHEETS E AUTENTICAÇÃO (MODO ONLINE)
; Instalado APENAS se não existir — preserva credenciais do usuário em atualizações
; ---------------------------------------------------------------
Source: "{#SourceDir}\credentials.json"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "{#SourceDir}\AuthNOC.json"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "{#SourceDir}\token_leitura.json"; DestDir: "{app}"; Flags: ignoreversion

; ---------------------------------------------------------------
; CACHE EACE (MODO OFFLINE / CACHE)
; Instalado APENAS se não existir — preserva cache mais recente do usuário
; ---------------------------------------------------------------
Source: "{#SourceDir}\eace_cache.json"; DestDir: "{app}"; Flags: onlyifdoesntexist

; ---------------------------------------------------------------
; PLANILHA EACE PADRÃO (GERADOR DE CONTRATO)
; Sempre substituída em atualizações para garantir versão padrão
; ---------------------------------------------------------------
Source: "{#SourceDir}\EACENOC.xlsx"; DestDir: "{app}"; Flags: ignoreversion

; NOTA: config_apps.json NÃO é incluído — o app gera automaticamente
; na primeira execução com os caminhos corretos do usuário atual.

[Icons]
; Menu Iniciar
Name: "{group}\{#MyAppName}";              Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Desinstalar {#MyAppName}";  Filename: "{uninstallexe}"

; Área de Trabalho (opcional — solicitado durante instalação)
Name: "{autodesktop}\{#MyAppName}";        Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Opção para iniciar o app após instalação
Filename: "{app}\{#MyAppExeName}"; Description: "Iniciar {#MyAppName} agora"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove a configuração do usuário ao desinstalar (caminhos salvos)
Type: files;     Name: "{app}\config_apps.json"
; Remove pasta se estiver vazia após desinstalação
Type: dirifempty; Name: "{app}"
