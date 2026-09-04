[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ConverterArguments
)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "First run: creating the virtual environment and installing dependencies..."
    & (Join-Path $projectRoot "install.ps1")
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

& $venvPython (Join-Path $projectRoot "native_pdf_to_md.py") @ConverterArguments
exit $LASTEXITCODE
