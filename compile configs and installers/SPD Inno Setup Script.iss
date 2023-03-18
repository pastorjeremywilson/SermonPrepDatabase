; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Sermon Prep Database"
#define MyAppVersion "3.3.7"
#define MyAppURL "https://sourceforge.net/projects/sermon-prep-database"
#define MyAppExeName "SermonPrepDatabase.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{15E4F012-5C1F-466C-B6AB-E2F33CA12348}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\dist\SermonPrepDatabase\gpl-3.0.rtf
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\output
OutputBaseFilename=Setup SPD v.{#MyAppVersion}
SetupIconFile=C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\dist\SermonPrepDatabase\resources\icons.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\dist\SermonPrepDatabase\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\pasto\Nextcloud\Documents\Python Workspace\Sermon Prep Database\compile configs and installers\dist\SermonPrepDatabase\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; WorkingDir: "{app}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

