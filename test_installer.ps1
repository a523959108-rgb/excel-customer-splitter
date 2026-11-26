# Run on a disposable Windows build runner. Installs and removes the current-user package.
$ErrorActionPreference = 'Stop'
$setup = (Resolve-Path "$PSScriptRoot\dist\installer\ExcelCustomerSplitter-Setup-x64.exe").Path
$root = Join-Path ([IO.Path]::GetTempPath()) ("Splitter-" + [guid]::NewGuid())
$program = Join-Path $root '软件 目录'
$data = Join-Path $root '客户 明细'
$db = Join-Path $data 'splitter.duckdb'
$exe = Join-Path $program 'Excel客户拆分.exe'
$profile = Join-Path $root 'profile'
$settings = Join-Path $profile 'ExcelCustomerSplitter\settings.json'
$originalLocal = $env:LOCALAPPDATA
function Invoke-Setup([string]$binary, [string[]]$options) {
    $proc = Start-Process $binary -ArgumentList $options -PassThru
    if (!$proc.WaitForExit(120000)) { $proc.Kill(); throw "Setup timed out." }
    if ($proc.ExitCode -ne 0) { throw "Setup failed with exit code $($proc.ExitCode)." }
}
function Start-AndCloseApp {
    try {
        $env:LOCALAPPDATA = $profile
        $app = Start-Process $exe -PassThru
        $readyBy = [datetime]::UtcNow.AddSeconds(45)
        do {
            Start-Sleep -Milliseconds 250
            $app.Refresh()
        } until (($app.MainWindowHandle -and (Test-Path $settings)) -or $app.HasExited -or [datetime]::UtcNow -ge $readyBy)
        if (!$app.MainWindowHandle -or !$app.CloseMainWindow() -or !$app.WaitForExit(20000)) { throw 'Installed app did not start and close normally.' }
        if ($app.ExitCode -ne 0) { throw 'Installed app exited with an error.' }
    } finally {
        if ($app -and !$app.HasExited) { $app.Kill(); $app.WaitForExit() }
        $env:LOCALAPPDATA = $originalLocal
    }
    if ((Get-Content -Encoding utf8 -Raw $settings | ConvertFrom-Json).database -ne $db) { throw 'Wrong active database.' }
}
# Each install gets the real wizard; only the clicks are suppressed on the build runner.
$quiet = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART')
Invoke-Setup $setup ($quiet + "/DIR=`"$program`"" + "/DBDIR=`"$data`"")
if (!(Test-Path $exe)) { throw 'Wrong application directory.' }
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
Remove-Item -LiteralPath $root -Recurse -Force
