[Setup]
AppId=ExcelCustomerSplitter
AppName=客户 Excel 拆分工具
AppVersion=1.0.1
DefaultDirName={userpf}\ExcelCustomerSplitter
DisableDirPage=no
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
SetupLogging=yes
UninstallDisplayIcon={app}\Excel客户拆分.exe
OutputDir=dist\installer
OutputBaseFilename=ExcelCustomerSplitter-Setup-x64
Compression=lzma2
SolidCompression=yes

[Languages]
Name: "zhcn"; MessagesFile: "{#ChineseMessages}"

[Files]
Source: "dist\Excel客户拆分\*"; DestDir: "{app}"; Excludes: "regions.json,customer_config.json"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\Excel客户拆分\regions.json"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "dist\Excel客户拆分\customer_config.json"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{tmp}\database-directory.txt"; DestDir: "{app}"; Flags: external ignoreversion

[Icons]
Name: "{userprograms}\客户 Excel 拆分工具"; Filename: "{app}\Excel客户拆分.exe"

[Run]
Filename: "{app}\Excel客户拆分.exe"; Description: "启动客户 Excel 拆分工具"; Flags: nowait postinstall skipifsilent

[Code]
var
  DatabasePage: TInputDirWizardPage;

procedure InitializeWizard;
var
  DefaultDir: String;
begin
  DatabasePage := CreateInputDirPage(wpSelectDir, '选择默认数据库目录',
    '客户明细数据库保存在什么位置？',
    '首次启动会在此目录打开或创建 splitter.duckdb。“新建库”也从此目录开始。' + #13#10 +
    '如果已经使用过本软件，启动时会继续打开上次使用的库。卸载软件会保留数据库。',
    False, '');
  DatabasePage.Add('默认数据库目录');
  DefaultDir := GetPreviousData('DatabaseDir', ExpandConstant('{localappdata}\ExcelCustomerSplitter'));
  DatabasePage.Values[0] := ExpandConstant('{param:DBDIR}');
  if DatabasePage.Values[0] = '' then DatabasePage.Values[0] := DefaultDir;
end;

procedure RegisterPreviousData(PreviousDataKey: Integer);
begin
  SetPreviousData(PreviousDataKey, 'DatabaseDir', DatabasePage.Values[0]);
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := MemoDirInfo + NewLine + NewLine + '默认数据库目录' + NewLine + Space + DatabasePage.Values[0];
end;

function CheckDirectory(const Folder, Caption: String): String;
var
  ProbeSource, ProbeDest: String;
  CreatedDir: Boolean;
begin
  if (Folder = '') or (ExtractFileDrive(Folder) = '') or
    (CompareText(ExpandFileName(Folder), Folder) <> 0) then begin
    Result := '请为' + Caption + '选择完整的文件夹路径。';
    Exit;
  end;
  ProbeSource := ExpandConstant('{tmp}\directory-check.tmp');
  if not SaveStringToFile(ProbeSource, 'test', False) then begin
    Result := '无法写入安装临时目录：' + ExpandConstant('{tmp}');
    Exit;
  end;
  Result := '当前账户无法在以下' + Caption + '创建或删除文件：' + #13#10 + Folder;
  CreatedDir := not DirExists(Folder);
  if CreatedDir and not ForceDirectories(Folder) then Exit;
  ProbeDest := AddBackslash(Folder) + ExtractFileName(ExpandConstant('{tmp}')) + '-write-test.tmp';
  if CopyFile(ProbeSource, ProbeDest, True) then
    if DeleteFile(ProbeDest) then Result := '';
  if CreatedDir then RemoveDir(Folder);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Problem, Suggested: String;
begin
  Result := True;
  if WizardSilent then Exit;
  if CurPageID = wpSelectDir then begin
    Problem := CheckDirectory(WizardDirValue, '软件目录');
    Suggested := ExpandConstant('{userpf}\ExcelCustomerSplitter');
  end else if CurPageID = DatabasePage.ID then begin
    Problem := CheckDirectory(Trim(DatabasePage.Values[0]), '数据库目录');
    Suggested := ExpandConstant('{localappdata}\ExcelCustomerSplitter');
  end else Exit;
  if Problem = '' then Exit;
  Result := False;
  Log(Problem);
  if MsgBox(Problem + #13#10#13#10 + '使用当前用户的默认目录？' + #13#10 +
      Suggested + #13#10#13#10 + '选“否”可自行更换位置。', mbConfirmation, MB_YESNO) = IDYES then begin
    if CurPageID = wpSelectDir then WizardForm.DirEdit.Text := Suggested
    else DatabasePage.Values[0] := Suggested;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  DataDir, ConfigFile: String;
  Lines: TArrayOfString;
begin
  Result := CheckDirectory(WizardDirValue, '软件目录');
  if Result <> '' then Exit;
  DataDir := Trim(DatabasePage.Values[0]);
  Result := CheckDirectory(DataDir, '数据库目录');
  if Result <> '' then Exit;
  if not ForceDirectories(DataDir) then begin
    Result := '无法创建数据库目录：' + DataDir;
    Exit;
  end;
  SetArrayLength(Lines, 1);
  Lines[0] := DataDir;
  ConfigFile := ExpandConstant('{tmp}\database-directory.txt');
  if not SaveStringsToUTF8File(ConfigFile, Lines, False) then
    Result := '无法保存默认数据库位置。';
end;
