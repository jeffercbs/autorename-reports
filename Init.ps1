$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$venvPath = Join-Path $scriptDir "venv"
$scriptPath = Join-Path $scriptDir "main.py"

if (-not (Test-Path $venvPath)) {
    Write-Output "Creating virtual environment..."
    python -m venv $venvPath
}

Write-Output "Create directory for saving files renamed..."
mkdir $scriptDir\renombrados -ErrorAction SilentlyContinue

Write-Output "validating if the excel file exists..."

if (-not (Test-Path $scriptDir\soportes.xlsx)) {
    Write-Output "The file soportes.xlsx does not exist, please add it to the root of the project"
    exit
}

Write-Output "Activating virtual environment..."
Write-Output "venvPath: $venvPath"
. $venvPath\Scripts\Activate.ps1

Write-Output "Installing dependencies..."
pip install -r $scriptDir\requirements.txt

Write-Output "Running script..."
python $scriptPath

Write-Output "Deactivating virtual environment..."
deactivate