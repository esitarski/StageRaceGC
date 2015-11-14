; -- StageRaceGC.iss --

[Setup]
#include "inno_setup.txt"
DefaultDirName={pf}\StageRaceGC
DefaultGroupName=StageRaceGC
UninstallDisplayIcon={app}\StageRaceGC.exe
Compression=lzma2/ultra64
SolidCompression=yes
SourceDir=dist\StageRaceGC
OutputDir=..\..\install
OutputBaseFilename=StageRaceGC_Setup
ChangesAssociations=yes

[Registry]
; Automatically configure StageRaceGC to launch .smr files.
Root: HKCR; Subkey: ".smr"; ValueType: string; ValueName: ""; ValueData: "StageRaceGC"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "StageRaceGC"; ValueType: string; ValueName: ""; ValueData: "StageRaceGC Race File"; Flags: uninsdeletekey
Root: HKCR; Subkey: "StageRaceGC\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\StageRaceGC.exe,0"
Root: HKCR; Subkey: "StageRaceGC\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\StageRaceGC.exe"" ""%1"""

[Tasks] 
Name: "desktopicon"; Description: "Create a &desktop icon"; 
	
[Files]
Source: "*.*"; DestDir: "{app}"; Flags: recursesubdirs

[Icons]
Name: "{group}\StageRaceGC"; Filename: "{app}\StageRaceGC.exe"
Name: "{userdesktop}\StageRaceGC"; Filename: "{app}\StageRaceGC.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\StageRaceGC.exe"; Description: "Launch StageRaceGC"; Flags: nowait postinstall skipifsilent
