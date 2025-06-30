;–––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; 						SAMS_Installer.iss
; Inno Setup script for Student Attendance Management System
;–––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Setup] Section: Installer metadata and basic behavior
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Setup]
; 1) Application name shown internally and in registry
AppName=Student Attendance Management System

; 2) Version string used internally (not shown to user)
AppVersion=1.0.0

; 3) Display name shown in Programs & Features and in installer window
AppVerName=Student Attendance Management System

; 4) Publisher name (your name or organization)
AppPublisher=Prashanth Kumar G

; 5) Default installation path under Program Files
DefaultDirName={pf}\Student Attendance Management System

; 6) File metadata: version shown in .exe properties
VersionInfoVersion=1.0.0

; 7) File metadata: displayed product version
VersionInfoProductVersion=1.0.0

; 8) Require administrator privileges (needed for OCX registration & SQL install)
PrivilegesRequired=admin

; 9) Skip the “Select Start Menu folder” page
DisableProgramGroupPage=yes

; 10) Ensure the application folder is created under Program Files
CreateAppDir=yes

; 11) Create a Icon for “Programs and Features” uninstall entry
UninstallDisplayIcon={app}\Student Attendance Management System.exe

; 12) Prevent setup from asking for restart at end
AlwaysRestart=no

; 13) Where to place the compiled Setup.exe (relative to this script)
OutputDir=.\Package

; 14) Base filename for the installer
OutputBaseFilename=SAMS_Setup

; 15) Use LZMA compression for smaller installer size
Compression=lzma
SolidCompression=yes

; 16) Icon to display on the setup .exe itself
SetupIconFile=Package\Images\sams.ico


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Languages] Section: UI language definitions
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Languages]
; 1) Define English as the installer’s UI language
Name: "english"; MessagesFile: "compiler:Default.isl"


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Dirs] Section: Create any needed folders with custom ACLs
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Dirs]
; Grant all users Modify rights on the Images folder so your app can save photos
Name: "{app}\Images"; Permissions: users-modify


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Files] Section: Files to include in the installer
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Files]
; 1) Copy compiled VB6 EXE into {app}
Source: "Package\Student Attendance Management System.exe"; DestDir: "{app}"; Flags: ignoreversion

; 2) Copy VB6 runtime and OCX dependencies into {app}\Bin
Source: "Package\Bin\msvbvm60.dll";    DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSCOMCTL.OCX";    DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSCOMCT2.OCX";   DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSCAL.OCX";      DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSCOMM32.OCX";   DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\COMDLG32.OCX";   DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSDATGRD.OCX";   DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSDBRPTR.DLL";   DestDir: "{app}\Bin"; Flags: ignoreversion
Source: "Package\Bin\MSSTDFMT.DLL";   DestDir: "{app}\Bin"; Flags: ignoreversion

; 3) Copy the offline SQL Express 2022 media into {tmp}\SQLExpressOffline
Source: "Package\SQLExpressOffline\*"; DestDir: "{tmp}\SQLExpressOffline"; Flags: ignoreversion recursesubdirs createallsubdirs

; 4) Copy the database backup file into {app}\Database
Source: "Package\Database\SAMS.bak"; DestDir: "{app}\Database"; Flags: ignoreversion

; 5) Copy the database restore script into {app}\Database
Source: "Package\Database\restore_sams.bat"; DestDir: "{app}\Database"; Flags: ignoreversion

; 6) Copy the Images folder (including all subfolders/files) into {app}\Images
Source: "Package\Images\*"; DestDir: "{app}\Images"; Flags: ignoreversion recursesubdirs

; 7) Copy the Documents folder (including all files) into {app}\Documents
Source: "Package\Documents\*"; DestDir: "{app}\Documents"; Flags: ignoreversion recursesubdirs


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Icons] Section: Shortcuts to create
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Icons]
; Create a desktop shortcut named “Student Attendance Management System”
Name: "{autodesktop}\Student Attendance Management System"; \
  Filename: "{app}\Student Attendance Management System.exe"; \
  WorkingDir: "{app}"

; Create Start Menu shortcut directly under All Programs (no subfolder)
Name: "{commonprograms}\Student Attendance Management System"; \
  Filename: "{app}\Student Attendance Management System.exe"; \
  WorkingDir: "{app}"


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Run] Section: Commands to execute during installation
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Run]
; 1) Register VB6 runtime (msvbvm60.dll—silent)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\msvbvm60.dll"""; \
  StatusMsg: "Registering VB6 runtime (msvbvm60.dll)..."; \
  Flags: runhidden

; 2) Register Common Controls (MSCOMCTL.OCX)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSCOMCTL.OCX"""; \
  StatusMsg: "Registering Common Controls (MSCOMCTL.OCX)..."; \
  Flags: runhidden

; 3) Register Common Controls 2 (MSCOMCT2.OCX)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSCOMCT2.OCX"""; \
  StatusMsg: "Registering Common Controls 2 (MSCOMCT2.OCX)..."; \
  Flags: runhidden

; 4) Register Calendar Control (MSCAL.OCX)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSCAL.OCX"""; \
  StatusMsg: "Registering Calendar Control (MSCAL.OCX)..."; \
  Flags: runhidden

; 5) Register Communications Control (MSCOMM32.OCX)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSCOMM32.OCX"""; \
  StatusMsg: "Registering Communications Control (MSCOMM32.OCX)..."; \
  Flags: runhidden

; 6) Register Common Dialog Control (COMDLG32.OCX)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\COMDLG32.OCX"""; \
  StatusMsg: "Registering Common Dialog Control (COMDLG32.OCX)..."; \
  Flags: runhidden

; 7) Register DataGrid Control (MSDATGRD.OCX)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSDATGRD.OCX"""; \
  StatusMsg: "Registering DataGrid Control (MSDATGRD.OCX)..."; \
  Flags: runhidden

; 8) Register Data Report runtime (MSDBRPTR.DLL)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSDBRPTR.DLL"""; \
  StatusMsg: "Registering Data Report runtime (MSDBRPTR.DLL)..."; \
  Flags: runhidden

; 9) Register Data Formatting Objects (MSSTDFMT.DLL)
Filename: "regsvr32"; \
  Parameters: "/s ""{app}\Bin\MSSTDFMT.DLL"""; \
  StatusMsg: "Registering Data Formatting Objects (MSSTDFMT.DLL)..."; \
  Flags: runhidden

; 10) Install SQL Server 2022 Express (engine or instance) if needed
Filename: "{tmp}\SQLExpressOffline\install_sql.bat"; \
  WorkingDir: "{tmp}\SQLExpressOffline"; \
  StatusMsg: "Installing SQL Server 2022 Express..."; \
  Flags: runhidden waituntilterminated; \
  Check: ShouldInstallSQL
  
; 11) Open credentials after successful installation
Filename: "{cmd}"; Parameters: "/C start """" ""{app}\Documents\DefaultLoginCredentials.txt"" & exit"; \
    Flags: postinstall runhidden nowait skipifsilent; \
    Description: "Show login credentials"; \
    StatusMsg: "Opening login credentials..."


;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
; [Code] Section: PascalScript functions for custom checks
;––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––;
[Code]
// Returns true if *any* SQL Server 2022 engine is installed
function IsSQLServer2022Installed(): Boolean;
begin
  Result := RegKeyExists(HKLM64, 'SOFTWARE\Microsoft\Microsoft SQL Server\160');
end;

// Returns true if instance SQLEXPRESS02 exists
function IsSQLExpress02Installed(): Boolean;
var
  ValueBuffer: String;
begin
  // We supply a variable (ValueBuffer) to receive the registry string.
  // RegQueryStringValue returns True if the value exists.
  Result := RegQueryStringValue(
    HKLM64,
    'SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
    'SQLEXPRESS02',
    ValueBuffer
  );
end;

// Only install SQL Express if either the engine is missing OR the instance is missing
function ShouldInstallSQL(): Boolean;
begin
  Result := (not IsSQLServer2022Installed()) or (not IsSQLExpress02Installed());
end;