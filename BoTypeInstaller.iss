[Setup]
AppName=BoType Word Add-in
AppVersion=2.0.0
AppPublisher=BoType
AppPublisherURL=https://github.com/bo-qian/BoType
DefaultDirName={autopf}\BoType
DefaultGroupName=BoType
OutputDir=.\
OutputBaseFilename=BoType_Setup_v2.0.0
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
DisableDirPage=no
DisableProgramGroupPage=yes
PrivilegesRequired=admin

[Files]
Source: "Publish\Application Files\BoType_2_0_0_0\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "BoType\BoType_TemporaryKey.cer"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: string; ValueName: "Description"; ValueData: "BoType Word Add-in"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: string; ValueName: "FriendlyName"; ValueData: "BoType Word Add-in"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\BoType.vsto|vstolocal"; Flags: uninsdeletekey

; Create the same registry keys in WOW6432Node for 32-bit Office running on 64-bit Windows.
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: string; ValueName: "Description"; ValueData: "BoType Word Add-in"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: string; ValueName: "FriendlyName"; ValueData: "BoType Word Add-in"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletevalue; Check: IsWin64
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\BoType"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\BoType.vsto|vstolocal"; Flags: uninsdeletekey; Check: IsWin64

[Run]
Filename: "certutil.exe"; Parameters: "-addstore -f ""Root"" ""{app}\BoType_TemporaryKey.cer"""; Flags: runhidden
Filename: "certutil.exe"; Parameters: "-addstore -f ""TrustedPublisher"" ""{app}\BoType_TemporaryKey.cer"""; Flags: runhidden

[UninstallRun]
Filename: "certutil.exe"; Parameters: "-delstore ""Root"" ""Bo Qian"""; Flags: runhidden; RunOnceId: "DelCert"
Filename: "certutil.exe"; Parameters: "-delstore ""TrustedPublisher"" ""Bo Qian"""; Flags: runhidden; RunOnceId: "DelCert2"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
