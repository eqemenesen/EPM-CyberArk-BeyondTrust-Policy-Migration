# Define the paths to the PowerShell scripts
$script1 = ".\script_AppGroup.ps1"
$script2 = ".\script_ContentGroups.ps1"
$script3 = ".\script_Workstyle.ps1"

# Define the base folder value
#$baseFolder = ".\Areas\FinansalProd"
#$baseFolder = ".\Areas\FinansalTest"
#$baseFolder = ".\Areas\GarantiProd"
$baseFolder = ".\Areas\GarantiTest"

# Run the scripts in order
try {
    Write-Host "Running Script1.ps1..." -ForegroundColor Green
    & $script1 -baseFolder $baseFolder

    Write-Host "Running Script2.ps1..." -ForegroundColor Green
    & $script2 -baseFolder $baseFolder

    Write-Host "Running Script3.ps1..." -ForegroundColor Green
    & $script3 -baseFolder $baseFolder

    Write-Host "All scripts executed successfully!" -ForegroundColor Green
} catch {
    Write-Error "An error occurred: $_"
}
