$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$VenvDir = Join-Path $ProjectRoot ".build-venv"
$BuildDir = Join-Path $ProjectRoot "build"
$DistDir = Join-Path $ProjectRoot "dist"
$ReleaseDir = Join-Path $ProjectRoot "release"
$ReleaseStage = Join-Path $ReleaseDir "FirewallSummary-Windows-x64"
$ReleaseZip = Join-Path $ReleaseDir "FirewallSummary-Windows-x64.zip"
$SelfTestResult = Join-Path $BuildDir "self-test.json"

function Find-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return @{
            Command = "py"
            Arguments = @("-3")
        }
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return @{
            Command = "python"
            Arguments = @()
        }
    }
    throw "Build computer needs Python 3.10+ to create the Windows package."
}

$Python = Find-PythonCommand
$PythonCommand = $Python.Command
$PythonArguments = @($Python.Arguments)

if (Test-Path $VenvDir) {
    Remove-Item $VenvDir -Recurse -Force
}
if (Test-Path $BuildDir) {
    Remove-Item $BuildDir -Recurse -Force
}
if (Test-Path $DistDir) {
    Remove-Item $DistDir -Recurse -Force
}
if (Test-Path $ReleaseStage) {
    Remove-Item $ReleaseStage -Recurse -Force
}
if (Test-Path $ReleaseZip) {
    Remove-Item $ReleaseZip -Force
}

& $PythonCommand @PythonArguments -c "import struct, sys; assert sys.version_info >= (3, 10), 'Python 3.10+ is required'; assert struct.calcsize('P') * 8 == 64, '64-bit Python is required'"
& $PythonCommand @PythonArguments -m venv $VenvDir

$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    throw "Failed to create the build virtual environment."
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r requirements-build.txt
& $VenvPython -m PyInstaller --clean --noconfirm firewall_summary.spec

$ExePath = Join-Path $DistDir "FirewallSummary.exe"
if (-not (Test-Path $ExePath)) {
    throw "PyInstaller did not create FirewallSummary.exe."
}

$SelfTestProcess = Start-Process `
    -FilePath $ExePath `
    -ArgumentList @("--self-test", "--self-test-output", "`"$SelfTestResult`"") `
    -Wait `
    -PassThru

if ($SelfTestProcess.ExitCode -ne 0) {
    throw "The packaged executable failed its runtime self-test."
}
if (-not (Test-Path $SelfTestResult)) {
    throw "The packaged executable did not create the self-test report."
}

$SelfTest = Get-Content $SelfTestResult -Raw | ConvertFrom-Json
if (-not $SelfTest.ok -or -not $SelfTest.frozen) {
    throw "The packaged executable is missing its embedded Python runtime or dependencies."
}

New-Item -ItemType Directory -Path $ReleaseStage -Force | Out-Null
Copy-Item $ExePath (Join-Path $ReleaseStage "FirewallSummary.exe")
Copy-Item (Join-Path $ProjectRoot "WINDOWS使用说明.txt") (Join-Path $ReleaseStage "WINDOWS使用说明.txt")
Copy-Item $SelfTestResult (Join-Path $ReleaseStage "构建自检.json")

New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null
Compress-Archive -Path (Join-Path $ReleaseStage "*") -DestinationPath $ReleaseZip -CompressionLevel Optimal

Write-Host ""
Write-Host "Windows release package created:"
Write-Host $ReleaseZip
Write-Host ""
Write-Host "The target computer does not need Python or any third-party packages."
