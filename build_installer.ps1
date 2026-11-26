$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$compiler = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
if (!(Test-Path $compiler)) {
    $compiler = (Get-Command ISCC.exe -ErrorAction Stop).Source
}
$buildDir = Join-Path $PSScriptRoot 'build\installer'
New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
$language = Join-Path $buildDir 'ChineseSimplified.isl'
$source = 'https://raw.githubusercontent.com/jrsoftware/issrc/99c68b09f759b86ab117560ceb89e29dcdb012ca/Files/Languages/Unofficial/ChineseSimplified.isl'
Invoke-WebRequest -Uri $source -OutFile $language
& $compiler "/DChineseMessages=$language" windows_installer.iss
if ($LASTEXITCODE) { throw 'Installer compilation failed.' }
