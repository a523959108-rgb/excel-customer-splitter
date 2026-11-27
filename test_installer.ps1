# Run on a disposable Windows build runner. Installs and removes the current-user package.
$ErrorActionPreference = 'Stop'
$setup = (Resolve-Path "$PSScriptRoot\dist\installer\ExcelCustomerSplitter-Setup-x64.exe").Path
$root = Join-Path ([IO.Path]::GetTempPath()) ("Splitter-" + [guid]::NewGuid())
$program = Join-Path $root '软件 目录'
$data = Join-Path $root '客户 明细'
$db = Join-Path $data 'splitter.duckdb'
$exe = Join-Path $program 'Excel客户拆分.exe'
$testProfile = Join-Path $root 'profile'
$settings = Join-Path $testProfile 'ExcelCustomerSplitter\settings.json'
$originalLocal = $env:LOCALAPPDATA
function Invoke-Setup([string]$binary, [string[]]$options) {
    $proc = Start-Process $binary -ArgumentList $options -PassThru
    if (!$proc.WaitForExit(120000)) { $proc.Kill(); throw "Setup timed out." }
    if ($proc.ExitCode -ne 0) { throw "Setup failed with exit code $($proc.ExitCode)." }
}
function Start-AndCloseApp {
    try {
        $env:LOCALAPPDATA = $testProfile
        $app = Start-Process $exe -PassThru
        $readyBy = [datetime]::UtcNow.AddSeconds(45)
        do {
            Start-Sleep -Milliseconds 250
            $app.Refresh()
            $ready = !$app.HasExited -and $app.MainWindowHandle.ToInt64() -ne 0 -and
                $app.MainWindowTitle -eq '客户 Excel 拆分工具' -and (Test-Path $settings)
        } until ($ready -or $app.HasExited -or [datetime]::UtcNow -ge $readyBy)
        if (!$ready) { throw "App did not open the expected window. Exit=$($app.HasExited), Title=$($app.MainWindowTitle), Settings=$(Test-Path $settings)" }
        if (!$app.CloseMainWindow() -or !$app.WaitForExit(20000)) { throw 'Could not close the installed application.' }
        if ($app.ExitCode -ne 0) { throw 'Installed app exited with an error.' }
    } finally {
        if ($app -and !$app.HasExited) { $app.Kill(); $app.WaitForExit() }
        $env:LOCALAPPDATA = $originalLocal
    }
    if ((Get-Content -Encoding utf8 -Raw $settings | ConvertFrom-Json).database -ne $db) { throw 'Wrong active database.' }
}
# Each install gets the real wizard; only the clicks are suppressed on the build runner.
$quiet = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CURRENTUSER')
# The build runner can normally write both locations. Deny file creation only in a test folder.
$blocked = New-Item -ItemType Directory -Path (Join-Path $root '只读目录') -Force
$acl = Get-Acl -LiteralPath $blocked.FullName
$deny = [Security.AccessControl.FileSystemAccessRule]::new(
    [Security.Principal.WindowsIdentity]::GetCurrent().User, 'CreateFiles', 'Deny')
$acl.AddAccessRule($deny)
Set-Acl -LiteralPath $blocked.FullName -AclObject $acl
foreach ($caption in @('软件目录', '数据库目录')) {
    $installDir = if ($caption -eq '软件目录') { $blocked.FullName } else { $program }
    $databaseDir = if ($caption -eq '数据库目录') { $blocked.FullName } else { $data }
    $logFile = Join-Path $root "$caption.log"
    $arguments = $quiet + "/DIR=`"$installDir`"" + "/DBDIR=`"$databaseDir`"" + "/LOG=`"$logFile`""
    $process = Start-Process $setup -ArgumentList $arguments -PassThru
    if (!$process.WaitForExit(30000)) { $process.Kill(); throw 'Folder permission check timed out.' }
    if ($process.ExitCode -ne 7 -or !(Select-String -LiteralPath $logFile -SimpleMatch "当前账户无法在以下$caption" -Quiet)) {
        throw "Setup did not identify the inaccessible folder: $caption (exit $($process.ExitCode))."
    }
    if (Test-Path $exe) { throw 'Setup copied the program before checking both folder permissions.' }
}
Invoke-Setup $setup ($quiet + "/DIR=`"$program`"" + "/DBDIR=`"$data`"")
if (!(Test-Path $exe)) { throw 'Wrong application directory.' }
$rule = New-NetFirewallRule -DisplayName ("Splitter offline test " + [guid]::NewGuid()) -Program $exe -Direction Outbound -Action Block -Profile Any
$previousPath = $env:Path
try {
    # Launch from outside the checkout with no Python, Office, or build tools on PATH.
    $env:Path = "$env:SystemRoot\System32;$env:SystemRoot"
    Push-Location $root
    try {
        Invoke-Setup $exe @('--self-test')
    } finally { Pop-Location }
} finally {
    $env:Path = $previousPath
    Remove-NetFirewallRule -Name $rule.Name
}
Write-Host 'Installed executable passed the offline Excel import, restart and export self-test.'

Start-AndCloseApp
& python -c "import duckdb,sys; c=duckdb.connect(sys.argv[1]); c.execute('CREATE TABLE install_probe AS SELECT 123 AS id'); c.close()" $db
if ($LASTEXITCODE) { throw 'Could not seed test database.' }
$newDefault = Join-Path $root '新的默认目录'
Invoke-Setup $setup ($quiet + "/DIR=`"$program`"" + "/DBDIR=`"$newDefault`"")
if ((Get-Content -Encoding utf8 -Raw (Join-Path $program 'database-directory.txt')).Trim() -ne $newDefault) { throw 'Upgrade did not save the new default.' }
Start-AndCloseApp
Invoke-Setup (Join-Path $program 'unins000.exe') $quiet
if (Test-Path $exe) { throw 'Uninstall left the executable behind.' }
if (!(Test-Path $db)) { throw 'Uninstall removed the database.' }
& python -c "import duckdb,sys; c=duckdb.connect(sys.argv[1],read_only=True); assert c.execute('SELECT id FROM install_probe').fetchall()==[(123,)]; c.close()" $db
if ($LASTEXITCODE) { throw 'Stored data changed during upgrade or uninstall.' }
Write-Host 'Verified custom program and database locations, first launch, upgrade and data retention after uninstall.'
$program = Join-Path $env:ProgramFiles ("ExcelCustomerSplitter-test-" + [guid]::NewGuid())
$exe = Join-Path $program 'Excel客户拆分.exe'
$allUsers = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/ALLUSERS')
$testProfile = Join-Path $root 'machine-install-user'
$settings = Join-Path $testProfile 'ExcelCustomerSplitter\settings.json'
$db = Join-Path $testProfile 'ExcelCustomerSplitter\splitter.duckdb'
Invoke-Setup $setup ($allUsers + "/DIR=`"$program`"")
try {
    if (!(Test-Path $exe) -or !(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ExcelCustomerSplitter_is1')) {
        throw 'Machine-wide installation in Program Files failed.'
    }
    Start-AndCloseApp
    $customDb = Join-Path $root "第二个盘的模拟目录 O'Brien"
    Invoke-Setup $setup ($allUsers + "/DIR=`"$program`"" + "/DBDIR=`"$customDb`"")
    if (!(Test-Path $customDb) -or (Get-Content -Encoding utf8 -Raw (Join-Path $program 'database-directory.txt')).Trim() -ne $customDb) {
        throw 'All-users installer did not prepare the selected custom database directory.'
    }
} finally {
    Invoke-Setup (Join-Path $program 'unins000.exe') @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART')
}
if (!(Test-Path $db)) { throw 'All-users uninstall removed the user database.' }
Write-Host 'Verified installation into Program Files, per-user database default and custom database location.'
Remove-Item -LiteralPath $root -Recurse -Force
