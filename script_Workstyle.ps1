param(
    [string]$baseFolder = ".\Areas\GarantiTest"
)

$useContent = $true

# -----------------------------------------------------------------------------
# Load Modules
# -----------------------------------------------------------------------------
try {
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -Force -ErrorAction Stop
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -Force -ErrorAction Stop
    Write-Host "Successfully imported Avecto modules."
} catch {
    Write-Host "ERROR: Failed to import Avecto modules. $_"
    exit 1
}

# -----------------------------------------------------------------------------
# Define Helper Functions
# -----------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------
# Define Paths & Counters
# -----------------------------------------------------------------------------
$reportFile         = Join-Path $baseFolder "GarantiMainPolicy.csv"  # Your CSV file
#$reportFile       = Join-Path $baseFolder "test.csv"
$logFilePath        = Join-Path $baseFolder "logs\logfile_Workstyle.log"
$outputFile         = Join-Path $baseFolder "generated_policy_Workstyles.xml"

if($useContent){
    $basePolicyXMLFile = Join-Path $baseFolder "generated_AppAndContentGroups.xml"
} else {
    $basePolicyXMLFile = Join-Path $baseFolder "generated_appGroup.xml"
}
    

Write-Host "Log file: $logFilePath"
Write-Host "Report file: $reportFile"

# Initialize Log File
"" | Out-File -FilePath $logFilePath -Force
Write-Log "Log file initialized." "INFO"

# -----------------------------------------------------------------------------
# Load Blank Policy Configuration
# -----------------------------------------------------------------------------
Write-Log "Loading blank policy configuration..." "INFO"
try {
    $PGConfig = Get-DefendpointSettings -LocalFile -FileLocation $basePolicyXMLFile -ErrorAction Stop
    Write-Log "Successfully loaded base policy configuration." "INFO"
} catch {
    Write-Log "Failed to load policy configuration. $_" "ERROR"
    throw
}

# -----------------------------------------------------------------------------
# Prepare Tracking for Existing Policies 
# -----------------------------------------------------------------------------
$existingPolicies = @($PGConfig.Policies | Select-Object -ExpandProperty Name)
Write-Log "Loaded existing policies: $($existingPolicies -join ', ')" "DEBUG"

# -----------------------------------------------------------------------------
# Gather WhiteList Apps from CSV
# -----------------------------------------------------------------------------
$WhiteListApps = @()
Write-Log "Parsing CSV for whitelisted applications..." "INFO"
try {
    Import-Csv -Path $reportFile -Delimiter "," | ForEach-Object {
        $PolicyName = $_."Policy Name"
        $PolicyType = $_."Policy Type"
        $Action = $_."Action"

        if ($PolicyType -eq "Predefined Application Group" -or $PolicyType -eq "Custom Application Group") {
            $WhiteListApps += @{Name=$PolicyName; Action=$Action}
        }
    }
    Write-Log "Completed gathering WhiteList Apps from CSV." "INFO"
} catch {
    Write-Log "Failed to parse CSV for WhiteList apps. $_" "ERROR"
    throw
}





# -----------------------------------------------------------------------------
# Ensure "White List" Policy Exists
# -----------------------------------------------------------------------------
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
    } catch {
        Write-Log "Failed to create 'White List' policy. $_" "ERROR"
        throw
    }
} else {
    Write-Log "'White List' policy already exists." "INFO"
}

# -----------------------------------------------------------------------------
# Process WhiteList Apps
# -----------------------------------------------------------------------------
Write-Log "Processing WhiteList Apps..." "INFO"

# Counters
$TotalWhiteListApps   = 0
$FailedWhiteListApps    = 0  # Apps with no matching AppGroup or overall logic error
$FailedWhiteListAppGroup  = 0  # Specifically failing to find the app group
$FailedWhiteListAssignments = 0  # Specifically failing to add assignment

# Get Message IDs
$allowMessage = $PGConfig.Messages | Where-Object { $_.Name -like "Allow Message (Elevate)" }
$blockMessage = $PGConfig.Messages | Where-Object { $_.Name -eq "Block Message" }
$allowBaloon  = $PGConfig.Messages | Where-Object { $_.Name -eq "Application Notification (Elevate)" }

$allowMessageId = $allowMessage.ID
$blockMessageId = $blockMessage.ID
$allowBaloonId  = $allowBaloon.ID

foreach ($appEntry in $WhiteListApps) {
    $appName = $appEntry.Name
    $appAction = $appEntry.Action
    $TotalWhiteListApps++
    try {
        $appGroup = $PGConfig.ApplicationGroups | Where-Object { $_.Name -eq $appName }

        if ($appGroup) {
            try {
                # Create an application assignment with dynamic Action and TokenType
                $appAssignment = New-Object Avecto.Defendpoint.Settings.ApplicationAssignment($PGConfig)

                if($appAction -eq "Elevate") {
                    $appAssignment.Action = "Allow"
                    $appAssignment.TokenType = "AddAdmin"
                } elseif ($appAction -eq "Block") {
                    $appAssignment.Action = "Block"
                    $appAssignment.TokenType = "Unmodified"

                } elseif ($appAction -eq "Run Normally") {
                    $appAssignment.Action = "Allow"
                    $appAssignment.TokenType = "Unmodified"

                }else {
                    $appAssignment.Action = "Allow"
                    $appAssignment.TokenType = "AddAdmin"
                }

                $appAssignment.Audit = "On"
                $appAssignment.PrivilegeMonitoring = "On"
                $appAssignment.ForwardBeyondInsight = $true
                $appAssignment.ForwardBeyondInsightReports = $true
                $appAssignment.ApplicationGroup = $appGroup

                # Add the application assignment to the "White List" policy
                $whiteListPolicy.ApplicationAssignments.Add($appAssignment)
                Write-Log "SUCCESS: Added '$appName' with action '$appAction' to the 'White List' policy." "INFO"
            }
            catch {
                $FailedWhiteListAssignments++
                Write-Log "EXCEPTION: Failed to add application assignment for '$appName'. $_" "ERROR"
            }
        } 
        else {
            $FailedWhiteListApps++
            $FailedWhiteListAppGroup++
            Write-Log "FAIL: No matching application group found for '$appName'." "ERROR"
            Write-Log "DETAILS: Application Group '$appName' not found in PGConfig.ApplicationGroups." "DEBUG"
        }
    } catch {
        $FailedWhiteListApps++
        Write-Log "EXCEPTION: Failed to process '$appName'. $_" "ERROR"
    }
}

# -----------------------------------------------------------------------------
# Process "Advanced Policy" Rows from CSV
# -----------------------------------------------------------------------------
$line = 1
$TotalAdvancedPolicies       = 0
$failedAdvancedPolicies      = 0  # Overall advanced policy failures (for reference)
$FailedPolicyAdd             = 0  # Specifically for failing to add the policy object
$FailedPolicyAppGroup        = 0  # Specifically for failing to add the ApplicationGroup assignment
$FailedPolicyContentGroup    = 0  # Specifically for failing to add the ContentGroup assignment
$FailedPolicyFilters         = 0  # Specifically for failing to add filters
$SkippedPolicies             = 0  # Count how many policies were skipped for any reason

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
        $Users              = $_."Users"
        $EndUserUI          = $_."End-User UI"

        # Only process "Advanced Policy"
        if ($PolicyType -ne "Advanced Policy") {
            return
        }

        # If CSV says it's inactive => skip with "SKIPPED" log
        if ($Active) {
            Write-Log "SKIPPED: Line $line => Policy '$PolicyName' is marked disabled (Active=No). Skipping creation." "WARN"
            $SkippedPolicies++
            return
        }

        # We only proceed if the PolicyName is valid for advanced policy
        if (-not $PolicyName) {
            Write-Log "SKIPPED: Line $line => Policy name is empty. Skipping creation." "WARN"
            $SkippedPolicies++
            return
        }

        # Common skip logic for mac-related or JIT
        if ($PolicyName -match "macOS" -or 
            $PolicyName -match "Default MAC Policy" -or
            $PolicyName -match "MAC OS" -or
            $PolicyName -clike "*JIT*") {

            Write-Log "SKIPPED: Line $line => Policy name '$PolicyName' is invalid/excluded (macOS, MAC OS, or JIT). Skipping." "WARN"
            $SkippedPolicies++
            return
        }

        if($Action -eq "Execute Script") {
            Write-Log "SKIPPED: Line $line => Policy name '$PolicyName' is invalid/excluded (Execute Script). Skipping." "WARN"
            $SkippedPolicies++
            return
        }

        $TotalAdvancedPolicies++

        # Check for duplicate policy name
        if ($existingPolicies -contains $PolicyName) {
            Write-Log "SKIPPED: Line $line => Policy '$PolicyName' already exists in PGConfig. Skipping creation." "WARN"
            $failedAdvancedPolicies++
            $SkippedPolicies++
            return
        }

        Write-Host "Processing line $line : $PolicyName"
        Write-Log "Processing line $line => policy '$PolicyName'." "INFO"

        # ----------------------------
        # Attempt to Create New Policy
        # ----------------------------
        $newPolicy = $null
        try {
            $newPolicy = New-Object Avecto.Defendpoint.Settings.Policy($PGConfig)
            $newPolicy.Name         = $PolicyName
            $newPolicy.Description  = $PolicyDescription
            $newPolicy.Disabled     = $false
            $newPolicy.GeneralRules = New-Object Avecto.Defendpoint.Settings.GeneralRules

            $newPolicy.GeneralRules.CaptureHostInfoRule.Configured     = "Enabled"
            $newPolicy.GeneralRules.CaptureUserInfoRule.Configured     = "Enabled"
            $newPolicy.GeneralRules.ProhibitAccountMgmtRule.Configured = "Enabled"

            Write-Log "Policy object '$PolicyName' created successfully." "INFO"
        }
        catch {
            Write-Log "EXCEPTION: Failed to create policy object for '$PolicyName'. $_" "ERROR"
            Write-Log "DETAILS: $($_.Exception.Message)" "DEBUG"
            $failedAdvancedPolicies++
            $FailedPolicyAdd++
            return
        }

        # ----------------------------
        # Application Rule
        # ----------------------------
        $AppGroupFound = $true
        try {
            $appGroup = $PGConfig.ApplicationGroups | Where-Object { $_.Name -eq $PolicyName }
            if ($null -eq $appGroup) {
                Write-Log "FAIL: No matching Application Group found for policy '$PolicyName'." "ERROR"
                Write-Log "DETAILS: Application Group '$PolicyName' not found in PGConfig.ApplicationGroups." "DEBUG"
                $FailedPolicyAppGroup++
                $AppGroupFound = $false
            }
            else {
                $appAssignment = New-Object Avecto.Defendpoint.Settings.ApplicationAssignment($PGConfig)

                # Translate CSV Action into Defendpoint fields
                switch ($Action) {
                    "Block" {
                        $appAssignment.Action    = "Block"
                        $appAssignment.TokenType = "Unmodified"
                    }
                    "Run Normally" {
                        $appAssignment.Action    = "Allow"
                        $appAssignment.TokenType = "Unmodified"
                    }
                    "Elevate" {
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

                # Handle End-User UI
                if ($EndUserUI -match "Elevate"){
                    $appAssignment.MessageIdAsText = $allowBaloonId
                    $appAssignment.ShowMessage     = $true
                    Write-Log "Elevate Notification Balloon added to '$PolicyName'." "INFO"
                } 
                elseif ($EndUserUI -match "Block"){
                    $appAssignment.MessageIdAsText = $blockMessageId
                    $appAssignment.ShowMessage     = $true
                    Write-Log "Block message added to '$PolicyName'." "INFO"
                } 
                elseif ($EndUserUI -match "Launch wi"){
                    $appAssignment.MessageIdAsText = $allowMessageId
                    $appAssignment.ShowMessage     = $true
                    Write-Log "Elevate Message added to '$PolicyName'." "INFO"
                }

                $appAssignment.ForwardBeyondInsight        = $true
                $appAssignment.ForwardBeyondInsightReports = $true

                # Add the application assignment to the policy
                $newPolicy.ApplicationAssignments.Add($appAssignment)

                # Add the on-demand application assignment
                #$newPolicy.ShellExtension.EnabledRunAs = $true
                #$newPolicy.ShellExtension.ApplicationAssignments.Add($appAssignment)

                Write-Log "SUCCESS: Application Assignment added to '$PolicyName'." "INFO"
            }
        }
        catch {
            Write-Log "EXCEPTION: Failed to add application assignment to '$PolicyName'. $_" "ERROR"
            Write-Log "DETAILS: $($_.Exception.Message)" "DEBUG"
            $FailedPolicyAppGroup++
            $AppGroupFound = $false
        }

        # ----------------------------
        # Content Rules
        # ----------------------------
        $ContentGroupFound = $true
        try {
            $contentGroup = $PGConfig.ContentGroups | Where-Object { $_.Name -eq $PolicyName }
            if ($contentGroup -ne $null) {
                $contentAssignment = New-Object Avecto.Defendpoint.Settings.ContentAssignment($PGConfig)
                $contentAssignment.ContentGroup = $contentGroup

                $contentAssignment.Action    = "Allow"
                $contentAssignment.TokenType = "AddAdmin"
                $contentAssignment.Audit     = "On"
                $contentAssignment.PrivilegeMonitoring = "On"

                $contentAssignment.ForwardBeyondInsight        = $true
                $contentAssignment.ForwardBeyondInsightReports = $true

                # Add the content assignment to the policy
                $newPolicy.ContentAssignments.Add($contentAssignment)
                Write-Log "SUCCESS: Content Assignment added to '$PolicyName'." "INFO"
                $ContentGroupFound = $true
            } 
            else {
                Write-Log "No matching Content Group found for '$PolicyName'." "Warn"
                Write-Log "DETAILS: Content Group '$PolicyName' not found in PGConfig.ContentGroups." "DEBUG"
                $ContentGroupFound = $false
            }
        }
        catch {
            Write-Log "EXCEPTION: Failed to add Content Assignment to '$PolicyName'. $_" "ERROR"
            Write-Log "DETAILS: $($_.Exception.Message)" "DEBUG"
            $FailedPolicyContentGroup++
            $ContentGroupFound = $false
        }

        # ----------------------------
        # Filters
        # ----------------------------
        $FiltersAdded = $false
        try {
            $newFilters = New-Object Avecto.Defendpoint.Settings.Filters
            $newFilters.FiltersLogic = "and"

            #----------
            # DeviceFilter
            #----------
            if (-not $AllComputers) {
                if ($SelectedComputers -or $SelectedComputers -ne "") {
                    $deviceFilter = New-Object Avecto.Defendpoint.Settings.DeviceFilter
                    $deviceFilter.Devices       = New-Object Avecto.Defendpoint.Settings.Devices
                    $deviceFilter.InverseFilter = $false

                    $selectedComputersList = $SelectedComputers -split ","
                    foreach ($computer in $selectedComputersList) {
                        $trimmedComputer = $computer.Trim()
                        # Simple heuristic to skip invalid entries
                        if ($trimmedComputer -match "^OU" -or $trimmedComputer -match "\\") {
                            Write-Log "Line $line : $PolicyName - Invalid SelectedComputers: $trimmedComputer. Skipping this entry." "WARN"
                            continue
                        }
                        else {
                            $deviceHostName = New-Object Avecto.Defendpoint.Settings.DeviceHostName
                            $deviceHostName.HostName = $trimmedComputer
                            $deviceFilter.Devices.DeviceHostNames.Add($deviceHostName)
                        }
                    }

                    # Only add the DeviceFilter if we have valid hostnames
                    if ($deviceFilter.Devices.DeviceHostNames.Count -gt 0) {
                        $newFilters.DeviceFilter = $deviceFilter
                        $FiltersAdded = $true
                        Write-Log "INFO: Device filters added to policy '$PolicyName'." "INFO"
                    }
                    else {
                        Write-Log "Line $line : $PolicyName - No valid hostnames found, skipping DeviceFilter." "WARN"
                    }
                }
                else {
                    Write-Log "Line $line : $PolicyName - No SelectedComputers provided, skipping DeviceFilter." "WARN"
                }
            }
            else {
                Write-Log "Line $line : $PolicyName - AllComputers is true, no DeviceFilter required." "INFO"
            }

            #----------
            # AccountsFilter
            #----------
            if ($Users -ne "") {
                try {
                    $accountsFilter = New-Object Avecto.Defendpoint.Settings.AccountsFilter
                    $accountsFilter.InverseFilter = $false
                    $accountsFilter.Accounts      = New-Object Avecto.Defendpoint.Settings.AccountList
                    $accountsFilter.Accounts.WindowsAccounts = New-Object System.Collections.Generic.List[Avecto.Defendpoint.Settings.Account]

                    $UsersList = $Users -split ","
                    foreach ($User in $UsersList) {
                        $trimmedUser = $User.Trim()

                        # Example skip logic
                        if ($trimmedUser -match "1") {
                            Write-Log "DETAILS: Skipping user (contains '1'): $trimmedUser for policy '$PolicyName'." "INFO"
                            continue
                        }
                        if (-not $trimmedUser) {
                            continue
                        }

                        $UserAccount = New-Object Avecto.Defendpoint.Settings.Account
                        if ($trimmedUser -like 'User *"' -or $trimmedUser -like 'Group *"') {
                            $UserName = ($trimmedUser -replace '^(User|Group)\s*"', '') -replace '"$', ''
                            if ($UserName -notmatch " ") {
                                $UserAccount.Name  = $UserName
                                $UserAccount.Group = ($trimmedUser -like 'Group *"')
                                $accountsFilter.Accounts.WindowsAccounts.Add($UserAccount)
                                Write-Log "DETAILS: Added user as Group='$($UserAccount.Group)': $UserName to policy '$PolicyName'." "INFO"
                            }
                            else {
                                Write-Log "DETAILS: Skipping user (contains space or invalid): $trimmedUser for '$PolicyName'." "INFO"
                            }
                        }
                        elseif ($trimmedUser -match "^.\\") {
                            $UserAccount.Name  = $trimmedUser
                            $UserAccount.Group = $false
                            $accountsFilter.Accounts.WindowsAccounts.Add($UserAccount)
                            Write-Log "DETAILS: Added user: $trimmedUser to policy '$PolicyName' as 'USER'." "INFO"
                        }
                        elseif ($trimmedUser -match "USR_SRV") {
                            $UserAccount.Name  = $trimmedUser
                            $UserAccount.Group = $true
                            $accountsFilter.Accounts.WindowsAccounts.Add($UserAccount)
                            Write-Log "DETAILS: Added user as GROUP: $trimmedUser to policy '$PolicyName' as 'GROUP'." "INFO"
                        }
                        else {
                            # Default assumption: treat it as a group
                            $UserAccount.Name  = $trimmedUser
                            $UserAccount.Group = $true
                            $accountsFilter.Accounts.WindowsAccounts.Add($UserAccount)
                            Write-Log "DETAILS: Added user/group: $trimmedUser to policy '$PolicyName' as 'GROUP'." "INFO"
                        }
                    }

                    if ($accountsFilter.Accounts.WindowsAccounts.Count -gt 0) {
                        $newFilters.AccountsFilter = $accountsFilter
                        $FiltersAdded = $true
                        Write-Log "AccountsFilter added to policy '$PolicyName' with accounts count: $($accountsFilter.Accounts.WindowsAccounts.Count)." "INFO"
                    }
                    else {
                        Write-Log "AccountsFilter is empty for '$PolicyName', skipping." "WARN"
                    }
                }
                catch {
                    Write-Log "Failed to build AccountsFilter for policy '$PolicyName'. Exception: $_" "ERROR"
                    $FailedPolicyFilters++
                }
            }

            #----------
            # Only assign filters if we have at least one of them
            #----------
            if ($FiltersAdded) {
                $newPolicy.Filters = $newFilters
                Write-Log "INFO: Filter(s) assigned to policy '$PolicyName'." "INFO"
            }
            else {
                Write-Log "INFO: No valid filters found for policy '$PolicyName', none assigned." "INFO"
            }
        }
        catch {
            Write-Log "EXCEPTION: Failed to create or assign filters for '$PolicyName'. $_" "ERROR"
            Write-Log "DETAILS: $($_.Exception.Message)" "DEBUG"
            $FailedPolicyFilters++
        }

        # ----------------------------
        # Add the new policy to the configuration
        # ----------------------------
        try {
            $PGConfig.Policies.Add($newPolicy)
            Write-Log "SUCCESS: Policy '$PolicyName' added to PGConfig." "INFO"
            # Add the policy name to the tracking list
            $existingPolicies += $PolicyName
        }
        catch {
            Write-Log "EXCEPTION: Failed to add policy '$PolicyName' to PGConfig. $_" "ERROR"
            Write-Log "DETAILS: $($_.Exception.Message)" "DEBUG"
            $failedAdvancedPolicies++
            $FailedPolicyAdd++
        }

        $line++
    }
}
catch {
    Write-Log "EXCEPTION: Failed while reading CSV lines. $_" "ERROR"
    throw
}

# -----------------------------------------------------------------------------
# SUMMARY REPORT
# -----------------------------------------------------------------------------
Write-Log "----- SUMMARY REPORT -----" "INFO"

# -- WhiteList --
Write-Log "WhiteList Apps Processing Complete." "INFO"
Write-Log "Total WhiteList Apps: $TotalWhiteListApps" "INFO"
Write-Log "Failed WhiteList Apps (No matching AppGroup overall): $FailedWhiteListApps" "INFO"
Write-Log " -- WhiteList:  Failed AppGroup lookups = $FailedWhiteListAppGroup" "INFO"
Write-Log " -- WhiteList:  Failed App Assignment    = $FailedWhiteListAssignments" "INFO"
if ($TotalWhiteListApps -gt 0) {
    $successCount = $TotalWhiteListApps - $FailedWhiteListApps
    $successRate  = [math]::Round(($successCount / $TotalWhiteListApps) * 100, 2)
    Write-Log "WhiteList Apps Success Rate: $successRate%" "INFO"
}
Write-Log "------------------------------------" "INFO"

# -- Advanced Policies --
Write-Log "Completed processing Advanced Policies." "INFO"
Write-Log "Total Advanced Policies (found in CSV): $TotalAdvancedPolicies" "INFO"
Write-Log "Skipped Policies (due to Active=No, macOS, JIT, etc.): $SkippedPolicies" "INFO"
Write-Log "Failed Advanced Policies (any type of error): $failedAdvancedPolicies" "INFO"

Write-Log "----- DETAILED FAILURES FOR ADVANCED POLICIES -----" "INFO"
Write-Log "Failed to create policy object:           $FailedPolicyAdd" "INFO"
Write-Log "Failed to add ApplicationGroup:           $FailedPolicyAppGroup" "INFO"
Write-Log "Failed to add ContentGroup:               $FailedPolicyContentGroup" "INFO"
Write-Log "Failed to add Filters (Device/Accounts):  $FailedPolicyFilters" "INFO"

if ($TotalAdvancedPolicies -gt 0) {
    $successCount = $TotalAdvancedPolicies - $failedAdvancedPolicies
    $successRate  = [math]::Round(($successCount / $TotalAdvancedPolicies) * 100, 2)
    Write-Log "Advanced Policy Overall Success Rate: $successRate%" "INFO"
}

# -----------------------------------------------------------------------------
# ADDITIONAL METRICS / PARTIAL FAIL RATES
# -----------------------------------------------------------------------------
Write-Log "===== ADDITIONAL METRICS FOR ADVANCED POLICIES =====" "INFO"

# How many policies were actually attempted (not skipped)?
$attemptedPolicies = $TotalAdvancedPolicies# - $SkippedPolicies
Write-Log "Policies actually attempted (excluding skipped): $attemptedPolicies" "INFO"

if ($attemptedPolicies -gt 0) {
    # Calculate fail rates for each category
    $policyCreationFailRate    = [math]::Round(($FailedPolicyAdd / $attemptedPolicies) * 100, 2)
    $appGroupFailRate          = [math]::Round(($FailedPolicyAppGroup / $attemptedPolicies) * 100, 2)
    $contentGroupFailRate      = [math]::Round(($FailedPolicyContentGroup / $attemptedPolicies) * 100, 2)
    $filterFailRate            = [math]::Round(($FailedPolicyFilters / $attemptedPolicies) * 100, 2)

    # You can also show success counts:
    $policyCreationSuccessCount   = $attemptedPolicies - $FailedPolicyAdd
    $appGroupSuccessCount         = $attemptedPolicies - $FailedPolicyAppGroup
    $contentGroupSuccessCount     = $attemptedPolicies - $FailedPolicyContentGroup
    $filterSuccessCount           = $attemptedPolicies - $FailedPolicyFilters

    Write-Log "Policy Creation Fail Rate:    $policyCreationFailRate%       (Failures: $FailedPolicyAdd, Success: $policyCreationSuccessCount)" "INFO"
    Write-Log "AppGroup Assignment Fail Rate:  $appGroupFailRate%           (Failures: $FailedPolicyAppGroup, Success: $appGroupSuccessCount)" "INFO"
    Write-Log "ContentGroup Assignment Fail Rate: $contentGroupFailRate%    (Failures: $FailedPolicyContentGroup, Success: $contentGroupSuccessCount)" "INFO"
    Write-Log "Filter Assignment Fail Rate:  $filterFailRate%               (Failures: $FailedPolicyFilters, Success: $filterSuccessCount)" "INFO"
}
else {
    Write-Log "No advanced policies were attempted (all skipped). Nothing to calculate fail rates for." "INFO"
}

# -----------------------------------------------------------------------------
# Continue with saving the config, etc.
# -----------------------------------------------------------------------------
Write-Log "Saving updated configuration to XML: $outputFile" "INFO"
try {
    Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation $outputFile -ErrorAction Stop
    Write-Log "SUCCESS: Policies with assignments saved to $outputFile." "INFO"
} catch {
    Write-Log "ERROR: Failed to save updated configuration. $_" "ERROR"
    throw
}

Write-Host "Done. See $logFilePath for details."
