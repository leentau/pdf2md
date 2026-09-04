[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$venvRoot = Join-Path $projectRoot ".venv"
$venvPython = Join-Path $venvRoot "Scripts\python.exe"
$bootstrapTemp = Join-Path $projectRoot ".bootstrap_tmp"

function Find-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
        if ($LASTEXITCODE -eq 0) {
            return @("py", "-3")
        }
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        & python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
        if ($LASTEXITCODE -eq 0) {
            return @("python")
        }
    }
    throw "Python 3.10 or newer is required. Install it from https://www.python.org/downloads/."
}

Set-Location -LiteralPath $projectRoot
$venvReady = $false
if (Test-Path -LiteralPath $venvPython) {
    $venvReady = Test-Path -LiteralPath (Join-Path $venvRoot "Scripts\pip.exe")
}

if (-not $venvReady) {
    if (Test-Path -LiteralPath $venvRoot) {
        $resolvedVenv = [System.IO.Path]::GetFullPath($venvRoot)
        $resolvedProject = [System.IO.Path]::GetFullPath($projectRoot)
        if (-not $resolvedVenv.StartsWith($resolvedProject + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to replace a virtual environment outside the project directory."
        }
        Remove-Item -LiteralPath $resolvedVenv -Recurse -Force
    }
    New-Item -ItemType Directory -Path $bootstrapTemp -Force | Out-Null
    $previousTemp = $env:TEMP
    $previousTmp = $env:TMP
    $env:TEMP = $bootstrapTemp
    $env:TMP = $bootstrapTemp
    $pythonCommand = Find-Python
    Write-Host "Creating Python virtual environment: $venvRoot"
    try {
        if ($pythonCommand.Count -eq 2) {
            & $pythonCommand[0] $pythonCommand[1] -m venv $venvRoot
        } else {
            & $pythonCommand[0] -m venv $venvRoot
        }
    } finally {
        $env:TEMP = $previousTemp
        $env:TMP = $previousTmp
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create the virtual environment."
    }
}

Write-Host "Installing/updating dependencies..."
New-Item -ItemType Directory -Path $bootstrapTemp -Force | Out-Null
$previousTemp = $env:TEMP
$previousTmp = $env:TMP
$env:TEMP = $bootstrapTemp
$env:TMP = $bootstrapTemp
try {
    & $venvPython -m pip install --disable-pip-version-check --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "Failed to upgrade pip." }
    & $venvPython -m pip install --disable-pip-version-check -r (Join-Path $projectRoot "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "Failed to install dependencies." }
} finally {
    $env:TEMP = $previousTemp
    $env:TMP = $previousTmp
}

& $venvPython -c "import docx, fitz, pypandoc, pypdf, requests, urllib3; print('Dependency check passed.')"
if ($LASTEXITCODE -ne 0) { throw "Dependency check failed." }
Remove-Item -LiteralPath $bootstrapTemp -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Installation completed."
Write-Host "Example: .\pdf2md.bat C:\path\manual.pdf --overwrite --report"
