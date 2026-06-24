[Setup]
AppId={{A7F3D2E1-4B8C-4F9A-B2D6-1E3A5C7F9B2D}
AppName=StationDeck
AppVersion=1.0.0
AppPublisher=StationDeck
DefaultDirName={autopf}\StationDeck
DefaultGroupName=StationDeck
DisableProgramGroupPage=yes
OutputDir=C:\Users\LENOVO\stationdeck\installer
OutputBaseFilename=StationDeck_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\Users\LENOVO\stationdeck\dist\StationDeck\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\StationDeck"; Filename: "{app}\StationDeck.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\StationDeck"; Filename: "{app}\StationDeck.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall StationDeck"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\StationDeck.exe"; Description: "Launch StationDeck now"; Flags: nowait postinstall skipifsilent; WorkingDir: "{app}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"