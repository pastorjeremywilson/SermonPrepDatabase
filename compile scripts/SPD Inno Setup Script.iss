; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Sermon Prep Database"
#define MyAppVersion "v.5.0.0"
#define MyAppURL "https://sourceforge.net/projects/sermon-prep-database"
#define MyAppExeName "SermonPrepDatabase.exe"
#define OutputLocation "C:\Users\pasto\Desktop\output\SermonPrepDatabase"
#define ResourceLocation "C:\Users\pasto\Desktop\output\SermonPrepDatabase\_internal\resources"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{15E4F012-5C1F-466C-B6AB-E2F33CA12348}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile={#ResourceLocation}\gpl-3.0.rtf
WizardImageFile={#ResourceLocation}\installImage.bmp
WizardSmallImageFile = {#ResourceLocation}\installImageSmall.bmp
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir="C:\Users\pasto\Desktop\output\SermonPrepDatabase\installer\"
OutputBaseFilename=Setup_SPD_{#MyAppVersion}
SetupIconFile={#ResourceLocation}\icons.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#OutputLocation}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#OutputLocation}\_internal\*"; DestDir: "{app}\_internal"; Flags: recursesubdirs createallsubdirs
Source: "C:\Users\pasto\Desktop\output\SermonPrepDatabase\_internal\README.html"; DestDir: "{app}\_internal"; Flags: isreadme
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; WorkingDir: "{app}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

