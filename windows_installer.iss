[Setup]
AppId=ExcelCustomerSplitter
AppName=客户 Excel 拆分工具
AppVersion=1.0.2
DefaultDirName={autopf}\ExcelCustomerSplitter
DisableDirPage=no
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UsePreviousPrivileges=no
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
Name: "{autoprograms}\客户 Excel 拆分工具"; Filename: "{app}\Excel客户拆分.exe"

[Run]
Filename: "{app}\Excel客户拆分.exe"; Description: "启动客户 Excel 拆分工具"; Flags: nowait postinstall skipifsilent

[Code]
const
  ProfileDatabaseDir = '%LOCALAPPDATA%\ExcelCustomerSplitter';

var
  DatabasePage: TInputDirWizardPage;
  UseProfileDir: TNewCheckBox;

procedure SetDatabaseControls(Sender: TObject);
begin
  DatabasePage.Edits[0].Visible := not UseProfileDir.Checked;
  DatabasePage.Buttons[0].Visible := not UseProfileDir.Checked;
  DatabasePage.PromptLabels[0].Visible := not UseProfileDir.Checked;
end;

function SelectedDatabaseDir: String;
begin
  if UseProfileDir.Checked then Result := ProfileDatabaseDir
  else Result := Trim(DatabasePage.Values[0]);
end;

procedure InitializeWizard;
var
  DefaultDir, CommandLineDir: String;
begin
  DatabasePage := CreateInputDirPage(wpSelectDir, '选择默认数据库目录',
    '客户明细数据库保存在什么位置？',
    '首次启动会在此目录打开或创建 splitter.duckdb。“新建库”也从此目录开始。' + #13#10 +
    '如果已经使用过本软件，启动时会继续打开上次使用的库。卸载软件会保留数据库。',
    False, '');
  DatabasePage.Add('默认数据库目录');
  if IsAdminInstallMode then DefaultDir := ProfileDatabaseDir
  else DefaultDir := ExpandConstant('{localappdata}\ExcelCustomerSplitter');
  DefaultDir := GetPreviousData('DatabaseDir', DefaultDir);
  CommandLineDir := ExpandConstant('{param:DBDIR}');
  if CommandLineDir <> '' then DefaultDir := CommandLineDir;
  UseProfileDir := TNewCheckBox.Create(DatabasePage);
  UseProfileDir.Parent := DatabasePage.Surface;
  UseProfileDir.Top := DatabasePage.Edits[0].Top + DatabasePage.Edits[0].Height + ScaleY(12);
  UseProfileDir.Width := DatabasePage.SurfaceWidth;
  UseProfileDir.Caption := '使用各用户自己的本地数据目录（%LOCALAPPDATA%\ExcelCustomerSplitter）';
  UseProfileDir.Visible := IsAdminInstallMode;
  UseProfileDir.Checked := IsAdminInstallMode and (DefaultDir = ProfileDatabaseDir);
  UseProfileDir.OnClick := @SetDatabaseControls;
  if DefaultDir = ProfileDatabaseDir then
    DatabasePage.Values[0] := ExpandConstant('{sd}\客户数据')
  else DatabasePage.Values[0] := DefaultDir;
  SetDatabaseControls(UseProfileDir);
end;

procedure RegisterPreviousData(PreviousDataKey: Integer);
begin
  SetPreviousData(PreviousDataKey, 'DatabaseDir', SelectedDatabaseDir);
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := MemoDirInfo + NewLine + NewLine + '默认数据库目录' + NewLine + Space + SelectedDatabaseDir;
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

function CheckDatabaseDir(const Folder: String): String;
var
  FolderCodeUnits, Command: String;
  I, ExitCode: Integer;
begin
  Result := '';
  if Folder = ProfileDatabaseDir then Exit;
  if not IsAdminInstallMode then begin
    Result := CheckDirectory(Folder, '数据库目录');
    Exit;
  end;
  if (ExtractFileDrive(Folder) = '') or (CompareText(ExpandFileName(Folder), Folder) <> 0)
    or (Pos('"', Folder) <> 0) then begin
    Result := '请为数据库目录选择完整的文件夹路径。';
    Exit;
  end;
  { Check a custom location as the user who launched Setup, including when UAC
    used a different administrator account. Create the directory as that user. }
  { Pass UTF-16 code units so quotes and other filename characters remain data. }
  FolderCodeUnits := IntToStr(Ord(Folder[1]));
  for I := 2 to Length(Folder) do
    FolderCodeUnits := FolderCodeUnits + ',' + IntToStr(Ord(Folder[I]));
  Command := '-NoProfile -NonInteractive -Command "try { $path=-join [char[]]@(' +
    FolderCodeUnits + '); $d=[IO.Directory]::CreateDirectory($path); ' +
    '$p=[IO.Path]::Combine($d.FullName,[IO.Path]::GetRandomFileName()); ' +
    '$f=[IO.FileStream]::new($p,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,' +
    '[IO.FileShare]::None,4096,[IO.FileOptions]::DeleteOnClose); $f.Dispose(); exit 0 } catch { exit 1 }"';
  if not ExecAsOriginalUser(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    Command, '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
    Result := '当前账户无法验证或写入以下数据库目录：' + #13#10 + Folder;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Problem, Suggested: String;
begin
  Result := True;
  if WizardSilent then Exit;
  if CurPageID = wpSelectDir then begin
    Problem := CheckDirectory(WizardDirValue, '软件目录');
    Suggested := ExpandConstant('{autopf}\ExcelCustomerSplitter');
    if (Problem <> '') and not IsAdminInstallMode then
      Problem := Problem + #13#10 + '若要安装到 Program Files，请重新启动安装包并选择“为所有用户安装”。';
  end else if CurPageID = DatabasePage.ID then begin
    Problem := CheckDatabaseDir(SelectedDatabaseDir);
    if IsAdminInstallMode then Suggested := ProfileDatabaseDir
    else Suggested := ExpandConstant('{localappdata}\ExcelCustomerSplitter');
  end else Exit;
  if Problem = '' then Exit;
  Result := False;
  Log(Problem);
  if MsgBox(Problem + #13#10#13#10 + '使用以下默认目录？' + #13#10 +
      Suggested + #13#10#13#10 + '选“否”可自行更换位置。', mbConfirmation, MB_YESNO) = IDYES then begin
    if CurPageID = wpSelectDir then WizardForm.DirEdit.Text := Suggested
    else if IsAdminInstallMode then begin
      UseProfileDir.Checked := True;
      DatabasePage.Values[0] := ExpandConstant('{sd}\客户数据');
      SetDatabaseControls(UseProfileDir);
    end else DatabasePage.Values[0] := Suggested;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  DataDir, ConfigFile: String;
  Lines: TArrayOfString;
begin
  Result := CheckDirectory(WizardDirValue, '软件目录');
  if Result <> '' then Exit;
  DataDir := SelectedDatabaseDir;
  Result := CheckDatabaseDir(DataDir);
  if Result <> '' then Exit;
  if DataDir = ProfileDatabaseDir then DataDir := ''
  else if not IsAdminInstallMode then begin
    if not ForceDirectories(DataDir) then begin
      Result := '无法创建数据库目录：' + DataDir;
      Exit;
    end;
  end;
  SetArrayLength(Lines, 1);
  Lines[0] := DataDir;
  ConfigFile := ExpandConstant('{tmp}\database-directory.txt');
  if not SaveStringsToUTF8File(ConfigFile, Lines, False) then
    Result := '无法保存默认数据库位置。';
end;
