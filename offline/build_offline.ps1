$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$PackageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Installer = Join-Path $PackageRoot "python-installer\python-3.12.10-amd64.exe"
$WheelDir = Join-Path $PackageRoot "wheels"
$SourceDir = Join-Path $PackageRoot "source"
$RuntimeDir = Join-Path $PackageRoot ".runtime"
$PythonDir = Join-Path $RuntimeDir "Python"
$PythonExe = Join-Path $PythonDir "python.exe"
$OutputDir = Join-Path $PackageRoot "output"
$BuildDir = Join-Path $RuntimeDir "build"
$DistDir = Join-Path $RuntimeDir "dist"
$SelfTestResult = Join-Path $RuntimeDir "self-test.json"

function Assert-File {
    param([string]$Path, [string]$Description)
    if (-not (Test-Path $Path -PathType Leaf)) {
        throw "Missing $Description`: $Path"
    }
}

Assert-File $Installer "Windows Python installer"
Assert-File (Join-Path $SourceDir "main.py") "application source"
Assert-File (Join-Path $SourceDir "firewall_summary.spec") "PyInstaller spec"
Assert-File (Join-Path $SourceDir "requirements-offline.txt") "offline requirements"

if (Test-Path $RuntimeDir) {
    Remove-Item $RuntimeDir -Recurse -Force
}
if (Test-Path $OutputDir) {
    Remove-Item $OutputDir -Recurse -Force
}

New-Item -ItemType Directory -Path $RuntimeDir -Force | Out-Null

Write-Host "[1/4] Installing the bundled private Python runtime..."
$InstallProcess = Start-Process `
    -FilePath $Installer `
    -ArgumentList @(
        "/quiet",
        "InstallAllUsers=0",
        "TargetDir=`"$PythonDir`"",
        "AssociateFiles=0",
        "Shortcuts=0",
        "Include_doc=0",
        "Include_debug=0",
        "Include_dev=1",
        "Include_exe=1",
        "Include_launcher=0",
        "Include_lib=1",
        "Include_pip=1",
        "Include_symbols=0",
        "Include_tcltk=1",
        "Include_test=0",
        "PrependPath=0"
    ) `
    -Wait `
    -PassThru

if ($InstallProcess.ExitCode -notin @(0, 3010)) {
    throw "Bundled Python installer failed with exit code $($InstallProcess.ExitCode)."
}
Assert-File $PythonExe "private Python runtime"

Write-Host "[2/4] Installing bundled wheels without network access..."
& $PythonExe -m pip install `
    --disable-pip-version-check `
    --no-index `
    --only-binary=:all: `
    --find-links $WheelDir `
    -r (Join-Path $SourceDir "requirements-offline.txt")

Write-Host "[3/4] Building the standalone Windows executable..."
Push-Location $SourceDir
try {
    & $PythonExe -m PyInstaller `
        --clean `
        --noconfirm `
        --workpath $BuildDir `
        --distpath $DistDir `
        firewall_summary.spec
}
finally {
    Pop-Location
}

$ExePath = Join-Path $DistDir "FirewallSummary.exe"
Assert-File $ExePath "packaged executable"

Write-Host "[4/4] Verifying the embedded runtime and dependencies..."
$SelfTestProcess = Start-Process `
    -FilePath $ExePath `
    -ArgumentList @("--self-test", "--self-test-output", "`"$SelfTestResult`"") `
    -Wait `
    -PassThru

if ($SelfTestProcess.ExitCode -ne 0) {
    throw "Packaged executable failed its self-test."
}
Assert-File $SelfTestResult "self-test report"

$SelfTest = Get-Content $SelfTestResult -Raw | ConvertFrom-Json
if (-not $SelfTest.ok -or -not $SelfTest.frozen) {
    throw "Packaged executable does not contain a working embedded Python runtime."
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
Copy-Item $ExePath (Join-Path $OutputDir "FirewallSummary.exe")
Copy-Item $SelfTestResult (Join-Path $OutputDir "build-self-test.json")
Copy-Item (Join-Path $SourceDir "WINDOWS_README.txt") (Join-Path $OutputDir "WINDOWS_README.txt")

Write-Host ""
Write-Host "Standalone executable created:"
Write-Host (Join-Path $OutputDir "FirewallSummary.exe")
Write-Host "The executable contains Python, Tk, openpyxl, and xlrd."
