# ------------------------------------
# Define Helper Functions
# ------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp [$Level] $Message"
    $logEntry | Out-File -FilePath $logFilePath -Append
    Write-Host $logEntry
}

# ------------------------------------
# Load Modules
# ------------------------------------
try {
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -Force -ErrorAction Stop
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -Force -ErrorAction Stop
    Write-Log "Successfully imported Avecto modules." "INFO"
}
catch {
    Write-Log "ERROR: Failed to import Avecto modules. $_" "ERROR"
    exit 1
}

# ------------------------------------
# Define Paths
# ------------------------------------
$baseFolder      = "."
$reportFile      = "$baseFolder\GarantiMainPolicy.csv"  # Your CSV file
$logFilePath     = "$baseFolder\logfile_Workstyle.txt"
$outputFile      = "$baseFolder\generated_policy_with_assignments.xml"
$AppCsvFilePath  = "./PolicySummary_Garanti.csv"  # If needed

Write-Host "Log file: $logFilePath"
Write-Host "Report file: $reportFile"

# Initialize Log File
"" | Out-File -FilePath $logFilePath -Force
Write-Log "Log file initialized." "INFO"

# ------------------------------------
# Load Blank Policy Configuration
# ------------------------------------
Write-Log "Loading blank policy configuration..." "INFO"
try {
    $PGConfig = Get-DefendpointSettings -LocalFile -FileLocation "$baseFolder\generated_appGroup.xml" -ErrorAction Stop
    Write-Log "Successfully loaded blank policy configuration." "INFO"
}
catch {
    Write-Log "ERROR: Failed to load policy configuration. $_" "ERROR"
    throw
}

# ------------------------------------
# Prepare Tracking for Existing Policies
# ------------------------------------
$existingPolicies = @($PGConfig.Policies | Select-Object -ExpandProperty Name)
Write-Log "Loaded existing policies: $($existingPolicies -join ', ')" "DEBUG"

# ------------------------------------
# Gather WhiteList Apps from CSV
# ------------------------------------
$WhiteListApps = @()
Write-Log "Parsing CSV for whitelisted applications..." "INFO"
try {
    Import-Csv -Path $reportFile -Delimiter "," | ForEach-Object {
        $PolicyName = $_."Policy Name"
        $PolicyType = $_."Policy Type"
        if ($PolicyType -eq "Predefined Application Group" -or $PolicyType -eq "Custom Application Group") {
            $WhiteListApps += $PolicyName
        }
    }
    Write-Log "Completed gathering WhiteList Apps from CSV." "INFO"
}
catch {
    Write-Log "ERROR: Failed to parse CSV for WhiteList apps. $_" "ERROR"
    throw
}

# ------------------------------------
# Ensure "White List" Policy Exists
# ------------------------------------
Write-Log "Checking for existing 'White List' policy..." "INFO"
$whiteListPolicy = $PGConfig.Policies | Where-Object { $_.Name -eq "White List" }
if (-not $whiteListPolicy) {
    Write-Log "No 'White List' policy found. Creating new one..." "INFO"
    try {
        $whiteListPolicy = New-Object Avecto.Defendpoint.Settings.Policy($PGConfig)
        $whiteListPolicy.Name        = "White List"
        $whiteListPolicy.Description = "Policy for whitelisted applications"
        $whiteListPolicy.Disabled    = $false
        $PGConfig.Policies.Add($whiteListPolicy)
        Write-Log "Created and added 'White List' policy." "INFO"
    }
    catch {
        Write-Log "ERROR: Failed to create 'White List' policy. $_" "ERROR"
        throw
    }
} else {
    Write-Log "'White List' policy already exists." "INFO"
}

# ------------------------------------
# Process WhiteList Apps
# ------------------------------------
Write-Log "Processing WhiteList Apps..." "INFO"
$TotalWhiteListApps   = 0
$FailedWhiteListApps  = 0

foreach ($appName in $WhiteListApps) {
    $TotalWhiteListApps++
    try {
        $appGroup = $PGConfig.ApplicationGroups | Where-Object { $_.Name -eq $appName }
        if ($appGroup) {
            # Create an application assignment
            $appAssignment = New-Object Avecto.Defendpoint.Settings.ApplicationAssignment($PGConfig)
            $appAssignment.Action              = "Allow"
            $appAssignment.TokenType           = "AddAdmin"
            $appAssignment.Audit               = "On"
            $appAssignment.PrivilegeMonitoring = "On"
            $appAssignment.ApplicationGroup    = $appGroup

            # Add the application assignment to the "White List" policy
            $whiteListPolicy.ApplicationAssignments.Add($appAssignment)
            Write-Log "SUCCESS: Added '$appName' to the 'White List' policy." "INFO"
        } else {
            $FailedWhiteListApps++
            Write-Log "FAIL: No matching application group found for '$appName'." "ERROR"

            # Detailed Logging for Missing Application Group
            Write-Log "DETAILS: Application Group '$appName' not found in PGConfig.ApplicationGroups." "DEBUG"
        }
    }
    catch {
        $FailedWhiteListApps++
        Write-Log "EXCEPTION: Failed to add '$appName'. $_" "ERROR"

        # Detailed Logging with Properties
        Write-Log "DETAILS: Application Assignment for '$appName' failed with error: $($_.Exception.Message)" "DEBUG"
    }
}

Write-Log "WhiteList Apps Processing Complete." "INFO"
Write-Log "Total WhiteList Apps: $TotalWhiteListApps" "INFO"
Write-Log "Failed WhiteList Apps: $FailedWhiteListApps" "INFO"
if ($TotalWhiteListApps -gt 0) {
    $successCount = $TotalWhiteListApps - $FailedWhiteListApps
    $successRate  = [math]::Round(($successCount / $TotalWhiteListApps) * 100, 2)
    Write-Log "WhiteList Apps Success Rate: $successRate%" "INFO"
}

Start-Sleep -Seconds 30 # Optional pause

# ------------------------------------
# Process "Advanced Policy" Rows from CSV
# ------------------------------------
$line = 0
$TotalAdvancedPolicies   = 0
$FailedAdvancedPolicies  = 0

Write-Log "Starting to process Advanced Policies from CSV..." "INFO"

try {
    Import-Csv -Path $reportFile -Delimiter "," | ForEach-Object {

        $PolicyType         = $_."Policy Type"
        $PolicyName         = $_."Policy Name"
        $PolicyDescription  = ($_."Policy Description" -replace "`r`n|`n|`r", " - ")
        $Active             = $_."Active" -eq "No"  # if "No" => disabled = $true
        $Action             = $_."Action"
        $SecurityToken      = $_."Security Token"
        $AllComputers       = $_."All Computers" -eq "Yes"
        $SelectedComputers  = $_."Selected Computers/Groups"

        # Only process "Advanced Policy"
        if ($PolicyType -ne "Advanced Policy") {
            return
        }

        # Increment total advanced policy attempts
        $TotalAdvancedPolicies++

        # Validate the policy name
        if (-not $PolicyName -or
           ($PolicyName -match "macOS" -or $PolicyName -match "Default MAC Policy" -or $PolicyName -match "MAC OS") -or
           $PolicyName -clike "*JIT*")
        {
            Write-Log "Line $line : Policy Name '$PolicyName' is invalid (e.g. 'macOS'). Skipping." "WARN"
            return
        }

        # Check for duplicate policy name
        if ($existingPolicies -contains $PolicyName) {
            Write-Log "Line $line : Policy '$PolicyName' already exists. Skipping." "WARN"
            $FailedAdvancedPolicies++
            return
        }

        Write-Host "Processing line $line : $PolicyName"
        Write-Log "INFO: Processing line $line for policy '$PolicyName'." "INFO"

        # Create a new policy in a try/catch
        try {
            $newPolicy = New-Object Avecto.Defendpoint.Settings.Policy($PGConfig)
            $newPolicy.Name        = $PolicyName
            $newPolicy.Description = $PolicyDescription
            $newPolicy.Disabled    = $Active  # if "No" => Disabled = $true

            # Find matching Application Group object
            $appGroup = $PGConfig.ApplicationGroups | Where-Object { $_.Name -eq $PolicyName }
            if ($appGroup -ne $null) {
                # Create an application assignment
                $appAssignment = New-Object Avecto.Defendpoint.Settings.ApplicationAssignment($PGConfig)

                switch ($Action) {
                    "Block" {
                        $appAssignment.Action    = "Block"
                        $appAssignment.TokenType = "None"
                    }
                    "Run Normally" {
                        $appAssignment.Action    = "Allow"
                        $appAssignment.TokenType = "AddAdmin"
                    }
                    default {
                        if ($SecurityToken -eq "Administrator") {
                            $appAssignment.Action    = "Allow"
                            $appAssignment.TokenType = "AddAdmin"
                        } else {
                            $appAssignment.Action    = "Allow"
                            $appAssignment.TokenType = "Unmodified"
                        }
                    }
                }

                $appAssignment.Audit               = "On"
                $appAssignment.PrivilegeMonitoring = "On"
                $appAssignment.ApplicationGroup    = $appGroup

                # Add the application assignment to the policy
                $newPolicy.ApplicationAssignments.Add($appAssignment)
                Write-Log "SUCCESS: Application Assignment added to '$PolicyName'." "INFO"
            } else {
                $FailedAdvancedPolicies++
                Write-Log "FAIL: No matching application group found for policy '$PolicyName'." "ERROR"

                # Detailed Logging for Missing Application Group
                Write-Log "DETAILS: Application Group '$PolicyName' not found in PGConfig.ApplicationGroups." "DEBUG"
            }

            # Handle Filters (AllComputers vs SelectedComputers)
            $newFilters = New-Object Avecto.Defendpoint.Settings.Filters
            $newFilters.FiltersLogic = "and"

            $DeviceFilter = New-Object Avecto.Defendpoint.Settings.DeviceFilter
            $DeviceFilter.Devices    = New-Object Avecto.Defendpoint.Settings.Devices
            $DeviceFilter.InverseFilter = $false

            if (-not $AllComputers) {
                if ($SelectedComputers -and $SelectedComputers -ne "") {
                    # Split and add each host
                    $selectedComputersList = $SelectedComputers -split ","
                    foreach ($computer in $selectedComputersList) {
                        $computer = $computer.Trim()
                        if ($computer) {
                            $deviceHostName = New-Object Avecto.Defendpoint.Settings.DeviceHostName
                            $deviceHostName.HostName = $computer
                            $DeviceFilter.Devices.DeviceHostNames.Add($deviceHostName)
                        }
                    }

                    if ($DeviceFilter.Devices.DeviceHostNames.Count -gt 0) {
                        $newFilters.DeviceFilter = $DeviceFilter
                        $newPolicy.Filters = $newFilters
                        Write-Log "INFO: Device filters added to policy '$PolicyName'." "INFO"
                    } else {
                        Write-Log "Line $line : $PolicyName - SelectedComputers is empty, no filters added." "WARN"
                    }
                } else {
                    Write-Log "Line $line : $PolicyName - No SelectedComputers provided, no filters added." "WARN"
                }
            } else {
                Write-Log "Line $line : $PolicyName - AllComputers is true, no filters required." "INFO"
            }

            # Assign to the new policy (if any filter logic is applicable)
            $newFilters.DeviceFilter = $DeviceFilter
            $newPolicy.Filters = $newFilters

            # Add the new policy to the configuration
            $PGConfig.Policies.Add($newPolicy)
            Write-Log "SUCCESS: Policy '$PolicyName' added to PGConfig." "INFO"

            # Add the policy name to the tracking list
            $existingPolicies += $PolicyName
        }
        catch {
            Write-Log "EXCEPTION: Failed to process policy '$PolicyName'. $_" "ERROR"

            # Detailed Logging with Properties
            Write-Log "DETAILS: Policy '$PolicyName' failed with error: $($_.Exception.Message)" "DEBUG"
            $FailedAdvancedPolicies++
        }

        # Increment line counter
        $line++
    }
}
catch {
    Write-Log "EXCEPTION: Failed while reading CSV lines. $_" "ERROR"
    throw
}


Write-Log "---------------------------------------" "INFO"
Write-Log "Completed processing Advanced Policies." "INFO"
Write-Log "Total Advanced Policies: $TotalAdvancedPolicies" "INFO"
Write-Log "Failed Advanced Policies: $FailedAdvancedPolicies" "INFO"
if ($TotalAdvancedPolicies -gt 0) {
    $successCount = $TotalAdvancedPolicies - $FailedAdvancedPolicies
    $successRate  = [math]::Round(($successCount / $TotalAdvancedPolicies) * 100, 2)
    Write-Log "Advanced Policy Success Rate: $successRate%" "INFO"
}

# ------------------------------------
# Save Updated Configuration
# ------------------------------------
Write-Log "Saving updated configuration to XML..." "INFO"
try {
    Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation $outputFile -ErrorAction Stop
    Write-Log "SUCCESS: Policies with application assignments saved to $outputFile." "INFO"
}
catch {
    Write-Log "ERROR: Failed to save updated configuration. $_" "ERROR"
    throw
}

Write-Host "Done. See $logFilePath for details."
