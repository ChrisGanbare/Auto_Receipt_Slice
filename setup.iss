; Inno Setup 安装脚本
; 用于创建专业的Windows安装程序
; 
; 使用方法：
; 1. 下载并安装 Inno Setup: https://jrsoftware.org/isdl.php
; 2. 打开 Inno Setup Compiler
; 3. 打开此文件 (setup.iss)
; 4. 点击 "Build" -> "Compile" 编译安装程序
; 5. 在 Output 文件夹中找到生成的安装程序

#define MyAppName "农行电子回单智能拆分工具"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "财务工具"
#define MyAppURL ""
#define MyAppExeName "农行电子回单智能拆分工具.exe"

[Setup]
; 注意：AppId的值用于标识应用程序，不要更改
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; 卸载时删除用户数据（可选，根据需要调整）
; UninstallDisplayIcon={app}\{#MyAppExeName}
LicenseFile=
InfoBeforeFile=
InfoAfterFile=
OutputDir=installer
OutputBaseFilename=农行电子回单智能拆分工具_安装程序
SetupIconFile=favicon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; 管理员权限（可选，根据需要调整）
PrivilegesRequired=lowest
; 使用默认英文界面（方案1）
; 如果需要中文界面，可以：
; 1. 下载 ChineseSimplified.isl 文件到项目目录
; 2. 取消注释下面的语言配置，并修改为：Name: "chinesesimp"; MessagesFile: "ChineseSimplified.isl"

[Languages]
; 使用默认英文界面，无需配置

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
Source: "dist\农行电子回单智能拆分工具.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
; 如果有其他需要包含的文件，在这里添加
; Source: "用户使用手册.pdf"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// 可选：添加自定义安装逻辑
// 例如：检查系统要求、创建配置文件等

