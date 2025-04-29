[Setup]
AppName=PDF和图片转Word 工具
AppVersion=2.0
DefaultDirName={pf}\PDFConverter
DefaultGroupName=PDFConverter
OutputDir=.
OutputBaseFilename=PDFConverter_Installer
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "build\PDFConverterInstaller\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PDFConverter"; Filename: "{app}\PDFConverter.exe"
Name: "{group}\卸载 PDFConverter"; Filename: "{uninstallexe}"
Name: "{commondesktop}\PDFConverter"; Filename: "{app}\PDFConverter.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务："

[Run]
Filename: "{app}\PDFConverter.exe"; Description: "运行 PDFConverter"; Flags: nowait postinstall skipifsilent
