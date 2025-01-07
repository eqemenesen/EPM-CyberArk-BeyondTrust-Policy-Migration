param(
    [string]$baseFolder = ".\Areas\GarantiTest"
)

### ---------------------------
### IMPORT NECESSARY MODULES
### ---------------------------
try {
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -Force -ErrorAction Stop
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -Force -ErrorAction Stop
    Write-Host "Successfully imported Avecto modules."
} catch {
    Write-Host "ERROR: Failed to import Avecto modules. $_"
    exit 1
}

### --------------------------
### SETUP & HELPER FUNCTIONS
### --------------------------

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )

    # We can add a date/time stamp or other prefix to each line for clarity
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp [$Level] $Message"
    $logEntry | Out-File -FilePath $logFilePath -Append
    Write-Host $logEntry
}

# This helper shortens any string to max $maxLen characters, appending "..."
# Policy Name must NOT be shortened.
function Compress-String {
    param(
        [string]$inputString,
        [int]$maxLen = 20
    )

    if ([string]::IsNullOrEmpty($inputString)) {
        return $inputString
    }
    if ($inputString.Length -le $maxLen) {
        return $inputString
    }
    else {
        return $inputString.Substring(0, $maxLen) + "..."
    }
}

### ---------------------------
### VARIABLES & INITIALIZATION
### ---------------------------

$reportFile     = "$baseFolder\PolicySummary_Garanti.csv"
$logFilePath    = "$baseFolder\logs\logfile_AppGroup.log"
$csvFilePath    = "$baseFolder\GarantiMainPolicy.csv"
$adminTasksFile = "$baseFolder\AdminTasks.csv"  # example path for your admin tasks file
$generatedXML   = "$baseFolder\generated_appGroup.xml"
$blankPolicyXML = "$baseFolder\blank_policy.xml"

Write-Host "Log: $logFilePath, Report: $reportFile"

# Clear old log file
"" | Out-File -FilePath $logFilePath -Force

Write-Log "Starting policy import and group creation."

### ------------------------------------------------------
### BUILD A POLICY DICTIONARY FOR TICKET NUMBER -> DESC
### ------------------------------------------------------

Write-Log "Importing ticket numbers from $csvFilePath"
$policyDictionary = @{}

try {
    Import-Csv -Path $csvFilePath | ForEach-Object {
        $policyName        = $_."Policy Name"
        $policyDescription = $_."Policy Description"
        
        if ($policyName) {
            $cleanedDescription = $policyDescription -replace "`r`n|`n|`r", " - "
            if ($cleanedDescription) {
                $policyDictionary[$policyName] = $cleanedDescription
            }
        }
    }
    Write-Log "Successfully built policy dictionary from $csvFilePath."
}
catch {
    Write-Log "Error building policy dictionary. Error: $($_.Exception.Message)" "ERROR"
    return
}

### ---------------------------
### LOAD DEFENDPOINT SETTINGS
### ---------------------------

Write-Log "Retrieving Defendpoint Settings from $blankPolicyXML"
try {
    $PGConfig = Get-DefendpointSettings -LocalFile -FileLocation $blankPolicyXML
    Write-Log "Loaded Defendpoint settings successfully."
}
catch {
    Write-Log "Failed to load $blankPolicyXML : $($_.Exception.Message)" "ERROR"
    return
}

### ---------------------------
### IMPORT ADMIN TASKS
### ---------------------------
Write-Log "Importing admin tasks from $adminTasksFile"
$adminTasks = $null
try {
    $headers = @("AdminPolicyName", "AdminApplicationName", "AdminType", "AdminPath", "AdminPathType", "AdminCLSID", "AdminCommandLine", "CMDMatchCase", "CMDMatchType", "AdminPublisher", "AdminPubType", "AdminPubMatchCase", "AdminProduct", "AdminProdMatchCase", "AdminProdType", "AdminServiceDisName")
    $adminTasks = Import-Csv -Path $adminTasksFile -Delimiter "," -Header $headers
    Write-Log "Admin Tasks: $adminTasks" "DEBUG"
}
catch {
    Write-Log "Failed to load Admin Tasks CSV: $($_.Exception.Message)" "ERROR"
    return
}

### ---------------------------
### MAIN LOOP: CREATE APP GROUPS
### ---------------------------

# We will keep counters for advanced logging
[int]$lineCount         = 1
[int]$addedCount        = 0
[int]$skippedCount      = 0
[int]$failedCount       = 0    # general failures
[int]$notSupportedCount = 0    # not-supported types
[int]$failedAppGroupCount = 0  # if a group can't be created (rare scenario)

# NEW: Track stats per ApplicationType
$typeStats = @{}  # e.g. { "Executable" = [PSCustomObject]@{Total=0; Added=0; Failed=0; Skipped=0; NotSupported=0} }

function Initialize-TypeStats {
    param([string]$appType)

    if (-not $typeStats.ContainsKey($appType)) {
        $typeStats[$appType] = [PSCustomObject]@{
            Total         = 0
            Added         = 0
            Skipped       = 0
            Failed        = 0
            NotSupported  = 0
        }
    }
}

Write-Log "Importing policy report from $reportFile"

$TargetAppGroupPre = "xxxxx"
$TargetAppGroup    = $null
$processedAppGroups = @{}

try {
    Import-Csv -Path $reportFile -Delimiter "," | ForEach-Object {

        $lineCount++

        # Grab all CSV properties into variables
        $Policy_Name                        = $_."Policy Name"
        $ApplicationType                    = $_."Application Type"
        $FileName                           = $_."File Name"
        $FileNameCompareAs                  = $_."File Name Compare As"
        $ChecksumAlgorithm                  = $_."Checksum Algorithm"
        $Checksum                           = $_."Checksum"
        $Owner                              = $_."Owner"
        $ArgumentsCompareAs                 = $_."Arguments Compare As"
        $SignedBy                           = $_."Signed By"
        $Publisher                          = $_."Publisher"
        $PublisherCompareAs                 = $_."Publisher Compare As"
        $ProductName                        = $_."Product Name"
        $ProductNameCompareAs               = $_."Product Name Compare As"
        $FileDescription                    = $_."File Description"
        $FileDescriptionCompareAs           = $_."File Description Compare As"
        $CompanyName                        = $_."Company Name"
        $CompanyNameCompareAs               = $_."Company Name Compare As"
        $ProductVersionFrom                 = $_."Product Version From"
        $ProductVersionTo                   = $_."Product Version To"
        $FileVersionFrom                    = $_."File Version From"
        $FileVersionTo                      = $_."File Version To"
        $ScriptFileName                     = $_."Script File Name"
        $ScriptShortcutName                 = $_."Script Shortcut Name"
        $MSPProductName                     = $_."MSI/MSP Product Name"
        $MSPProductNameCompareAs            = $_."MSI/MSP Product Name Compare As"
        $MSPCompanyName                     = $_."MSI/MSP Company Name"
        $MSPCompanyNameCompareAs            = $_."MSI/MSP Company Name Compare As"
        $MSPProductVersionFrom              = $_."MSI/MSP Product Version From"
        $MSPProductVersionTo                = $_."MSI/MSP Product Version To"
        $MSPProductCode                     = $_."MSI/MSP Product Code"
        $MSPUpgradeCode                     = $_."MSI/MSP Upgrade Code"
        $WebApplicationURL                  = $_."Web Application URL"
        $WebApplicationURLCompareAs         = $_."Web Application URL Compare As"
        $WindowsAdminTask                   = $_."Windows Admin Task"
        $AllActiveXInstallationAllowed      = $_."All ActiveX Installation Allowed"
        $ActiveXSourceURL                   = $_."ActiveX Source URL"
        $ActiveXSourceCompareAs             = $_."ActiveX Source Compare As"
        $ActiveXMIMEType                    = $_."ActiveX MIME Type"
        $ActiveXMIMETypeCompareAs           = $_."ActiveX MIME Type Compare As"
        $ActiveXCLSID                       = $_."ActiveX CLSID"
        $ActiveXVersionFrom                 = $_."ActiveX Version From"
        $ActiveXVersionTo                   = $_."ActiveX Version To"
        $FileFolderPath                     = $_."File/Folder Path"
        $FileFolderType                     = $_."File/Folder Type"
        $FolderWithsubfoldersandfiles       = $_."Folder With subfolders and files"
        $RegistryKey                        = $_."Registry Key"
        $COMDisplayName                     = $_."COM Display Name"
        $COMCLSID                           = $_."COM CLSID"
        $ServiceName                        = $_."Service Name"
        $ElevatePrivilegesforChildProcesses = $_."Elevate Privileges for Child Processes"
        $RemoveAdminrightsfromFileOpen      = $_."Remove Admin rights from File Open/Save common dialogs"
        $ApplicationDescription             = $_."Application Description"

        Write-Log "Processing line: $lineCount, Policy Name: $Policy_Name" "INFO"
        Write-Log " CSV Values => ApplicationType: $ApplicationType, FileName: $FileName, ChecksumAlg: $ChecksumAlgorithm, Checksum: $Checksum, SignedBy: $SignedBy, Publisher: $Publisher, ProductName: $ProductName, FileDescription: $FileDescription, ServiceName: $ServiceName" "DEBUG"
        Write-Log " CSV Values => ProductVersionFrom: $ProductVersionFrom, ProductVersionTo: $ProductVersionTo, FileVersionFrom: $FileVersionFrom, FileVersionTo: $FileVersionTo, WindowsAdminTask: $WindowsAdminTask" "DEBUG"

        try {
            # Skip known excluded lines
            if ([string]::IsNullOrEmpty($Policy_Name) -or 
                $Policy_Name -match "macOS" -or 
                $Policy_Name -match "Default MAC Policy" -or 
                $Policy_Name -clike 'Usage of `"JIT*' -or
                $ApplicationType -like "Script") {
                
                Write-Log "Line $lineCount - Skipped Policy '$Policy_Name' (excluded pattern)." "WARN"
                $skippedCount++

                # Also track it per-ApplicationType, if it exists
                if ($ApplicationType) {
                    Initialize-TypeStats -appType $ApplicationType
                    $typeStats[$ApplicationType].Total++
                    $typeStats[$ApplicationType].Skipped++
                }
                return
            }

            # Initialize statistics for this ApplicationType
            Initialize-TypeStats -appType $ApplicationType
            $typeStats[$ApplicationType].Total++

            # Attempt to handle group creation or reuse
            try {
                if (-not $processedAppGroups.ContainsKey($Policy_Name)) {
                    if ($Policy_Name -ne $TargetAppGroupPre) {
                        # If there is an existing group, add it to the PGConfig
                        if ($TargetAppGroup) {
                            $PGConfig.ApplicationGroups.Add($TargetAppGroup)
                            Write-Log "Added AppGroup '$($TargetAppGroup.Name)' to PGConfig." "DEBUG"
                        }

                        # Create a new group
                        $TargetAppGroup = New-Object Avecto.Defendpoint.Settings.ApplicationGroup
                        $TargetAppGroup.Name        = $Policy_Name
                        $TargetAppGroup.Description = $Policy_Name
                        $TargetAppGroupPre          = $Policy_Name
                        
                        # Mark this app group as processed
                        $processedAppGroups[$Policy_Name] = $true
                        Write-Log "Created new application group '$Policy_Name'." "INFO"
                    }
                }
            }
            catch {
                # If group creation fails
                $failedAppGroupCount++
                throw $_  # rethrow so the outer catch captures it
            }

            # Create a new application object
            $PGApp = New-Object Avecto.Defendpoint.Settings.Application $PGConfig
            Write-Log "Created new Avecto application object for policy '$Policy_Name'." "DEBUG"

            # If we have a custom description
            if ($ApplicationDescription) {
                $PGApp.Description = $ApplicationDescription
            }
            elseif ($FileName) {
                $PGApp.Description = $FileName
            }

            # Determine Application Type
            switch -Wildcard ($ApplicationType) {
                "ActiveX Control" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ActiveX
                }
                "Admin Tasks" {
                    $AdminApplicationName = $WindowsAdminTask
                    $adminTask = $adminTasks | Where-Object { $_.AdminApplicationName -eq $AdminApplicationName }
                    if ("" -eq $adminTask -or $null -eq $adminTask) {
                        Write-Log "[AdminTask] Line $lineCount - Admin Task '$AdminApplicationName' not found in adminTasksFile." "WARN"
                        $failedCount++
                        $typeStats[$ApplicationType].Failed++
                        return
                    }

                    switch ($adminTask.AdminType) {
                        "Executable" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Executable
                            $PGApp.AppxPackageNameMatchCase = "Contains"
                            $PGApp.FileName     = $adminTask.AdminPath
                            $PGApp.ProductName  = $adminTask.AdminApplicationName
                            $PGApp.DisplayName  = $adminTask.AdminApplicationName
                            if ($adminTask.AdminCommandLine -ne $null) {
                                $PGApp.CmdLine = $adminTask.AdminCommandLine
                                $PGApp.CmdLineMatchCase = "false"
                                $PGApp.CmdLineStringMatchType = "Contains"
                            }
                            if($adminTask.AdminPublisher -ne $null) {
                                $PGApp.Publisher = $adminTask.AdminPublisher
                                $PGApp.CheckPublisher = $true
                                $PGApp.PublisherStringMatchType = "Contains"
                                $PGApp.PublisherMatchCase = $true
                            }
                        }
                        "COMClass" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::COMClass
                            $PGApp.DisplayName = $adminTask.AdminApplicationName
                            $PGApp.CheckAppID  = "true"
                            $PGApp.CheckCLSID  = "true"
                            $PGApp.AppID       = $adminTask.AdminCLSID
                            $PGApp.CLSID       = $adminTask.AdminCLSID
                        }
                        "ManagementConsoleSnapin" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ManagementConsoleSnapin
                            $PGApp.DisplayName = $adminTask.AdminApplicationName
                            $PGApp.AppxPackageNameMatchCase = "Contains"
                            $PGApp.FileName = $adminTask.AdminPath
                            if ($adminTask.AdminCommandLine -ne $null) {
                                $PGApp.CmdLine = $adminTask.AdminCommandLine
                                $PGApp.CmdLineMatchCase = "false"
                                $PGApp.CmdLineStringMatchType = "Contains"
                            }
                            if($adminTask.AdminPublisher -ne $null) {
                                $PGApp.Publisher = $adminTask.AdminPublisher
                                $PGApp.CheckPublisher = $true
                                $PGApp.PublisherStringMatchType = "Contains"
                                $PGApp.PublisherMatchCase = $true
                            }
                        }
                        "ControlPanelApplet" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ControlPanelApplet
                            $PGApp.CheckAppID  = "true"
                            $PGApp.CheckCLSID  = "true"
                            $PGApp.AppID       = $adminTask.AdminCLSID
                            $PGApp.CLSID       = $adminTask.AdminCLSID
                        }
                        "Service" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Service
                            $PGApp.ServiceName                = $adminTask.AdminPath
                            $PGApp.CheckServiceName           = $true
                            $PGApp.ServiceNamePatternMatching = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                            $PGApp.ServicePause               = $true
                            $PGApp.ServiceStart               = $true
                            $PGApp.ServiceStop                = $true
                            $PGApp.ServiceConfigure           = $true
                            $PGApp.ServiceDisplayName         = $adminTask.AdminServiceDisName
                            $PGApp.CheckServiceDisplayName    = $true
                            $PGApp.ServiceDisplayNamePatternMatching = $true
                        }
                        default {
                            Write-Log "[AdminTask] Line $lineCount - Policy '$Policy_Name', AdminType '$($adminTask.AdminType)' is not supported." "WARN"
                            $failedCount++
                            $typeStats[$ApplicationType].NotSupported++
                            return
                        }
                    }
                }
                "Executable" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Executable
                }
                "COM Object" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::COMClass
                }
                "MSI/MSP Installation" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::InstallerPackage
                }
                "Script" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::BatchFile
                }
                "Dynamic-Link Library" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', 'Dynamic-Link Library' not supported." "WARN"
                    $notSupportedCount++
                    $typeStats[$ApplicationType].NotSupported++
                    return
                }
                "Service" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Service
                }
                "Registry Key" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::RegistrySettings
                }
                "File or Directory System Entry" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' will be added later." "WARN"
                    $skippedCount++
                    $typeStats[$ApplicationType].Skipped++
                    return
                }
                "Microsoft Update (MSU)" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::InstallerPackage
                }
                "Web Application" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is NOT supported (error)." "ERROR"
                    $failedCount++
                    $notSupportedCount++
                    $typeStats[$ApplicationType].NotSupported++
                    return
                }
                "Win App" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is NOT supported (error)." "ERROR"
                    $failedCount++
                    $notSupportedCount++
                    $typeStats[$ApplicationType].NotSupported++
                    return
                }
                default {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is UNRECOGNIZED (error)." "ERROR"
                    $failedCount++
                    $notSupportedCount++
                    $typeStats[$ApplicationType].NotSupported++
                    return
                }
            }

            # File Name
            if ($FileName) {
                Write-Log "Setting FileName property for Policy '$Policy_Name' to '$FileName'." "DEBUG"
                $PGApp.CheckFileName = $true
                if ($FileName.StartsWith("*")) {
                    $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                    $PGApp.FileName = $FileName.Substring(1)
                }
                elseif ($FileName.EndsWith("*")) {
                    $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                    $PGApp.FileName = $FileName.Substring(0, $FileName.Length - 1)
                }
                else {
                    $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                    $PGApp.FileName = $FileName
                }
            }

            # CheckSum
            if ($ChecksumAlgorithm -eq "SHA1" -and $Checksum) {
                Write-Log "Setting CheckFileHash for Policy '$Policy_Name' to '$Checksum'." "DEBUG"
                $PGApp.CheckFileHash = $true
                $PGApp.FileHash      = $Checksum
            }

            # Publisher
            if ($SignedBy -eq "Specific Publishers" -and $Publisher) {
                Write-Log "Setting Publisher for Policy '$Policy_Name' to '$Publisher'." "DEBUG"
                $PGApp.CheckPublisher = $true
                if ($Publisher.StartsWith("*")) {
                    $PGApp.PublisherStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                    $PGApp.Publisher = $Publisher.Substring(1)
                }
                elseif ($Publisher.EndsWith("*")) {
                    $PGApp.PublisherStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                    $PGApp.Publisher = $Publisher.Substring(0, $Publisher.Length - 1)
                }
                else {
                    $PGApp.PublisherStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                    $PGApp.Publisher = $Publisher
                }
            }

            # Product Name
            if ($ProductName) {
                Write-Log "Setting ProductName for Policy '$Policy_Name' to '$ProductName'." "DEBUG"
                $PGApp.CheckProductName = $true
                if ($ProductName.StartsWith("*")) {
                    $PGApp.ProductNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                    $PGApp.ProductName = $ProductName.Substring(1)
                }
                elseif ($ProductName.EndsWith("*")) {
                    $PGApp.ProductNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                    $PGApp.ProductName = $ProductName.Substring(0, $ProductName.Length - 1)
                }
                else {
                    $PGApp.ProductNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                    $PGApp.ProductName = $ProductName
                }
            }

            # File Description
            if ($FileDescription) {
                Write-Log "Setting FileDescription for Policy '$Policy_Name' to '$FileDescription'." "DEBUG"
                $PGApp.CheckProductDesc = $true
                if ($FileDescription.StartsWith("*")) {
                    $PGApp.ProductDescStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                    $PGApp.ProductDesc = $FileDescription.Substring(1)
                }
                elseif ($FileDescription.EndsWith("*")) {
                    $PGApp.ProductDescStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                    $PGApp.ProductDesc = $FileDescription.Substring(0, $FileDescription.Length - 1)
                }
                else {
                    $PGApp.ProductDescStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                    $PGApp.ProductDesc = $FileDescription
                }
            }

            # Product Version From/To
            if ($ProductVersionFrom) {
                Write-Log "Setting MinProductVersion for Policy '$Policy_Name' to '$ProductVersionFrom'." "DEBUG"
                $PGApp.CheckMinProductVersion = $true
                $PGApp.MinProductVersion      = $ProductVersionFrom
            }
            if ($ProductVersionTo) {
                Write-Log "Setting MaxProductVersion for Policy '$Policy_Name' to '$ProductVersionTo'." "DEBUG"
                $PGApp.CheckMaxProductVersion = $true
                $PGApp.MaxProductVersion      = $ProductVersionTo
            }

            # File Version From/To
            if ($FileVersionFrom) {
                Write-Log "Setting MinFileVersion for Policy '$Policy_Name' to '$FileVersionFrom'." "DEBUG"
                $PGApp.CheckMinFileVersion = $true
                $PGApp.MinFileVersion      = $FileVersionFrom
            }
            if ($FileVersionTo) {
                Write-Log "Setting MaxFileVersion for Policy '$Policy_Name' to '$FileVersionTo'." "DEBUG"
                $PGApp.CheckMaxFileVersion = $true
                $PGApp.MaxFileVersion      = $FileVersionTo
            }

            # Service
            if ($ServiceName) {
                Write-Log "Setting ServiceName for Policy '$Policy_Name' to '$ServiceName'." "DEBUG"
                $PGApp.CheckServiceName           = $true
                $PGApp.ServiceNamePatternMatching = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                $PGApp.ServiceName                = $ServiceName
                $PGApp.ServicePause               = $true
                $PGApp.ServiceStart               = $true
                $PGApp.ServiceStop                = $true
                $PGApp.ServiceConfigure           = $true
            }

            # Elevate / Remove Admin rights
            if ($ElevatePrivilegesforChildProcesses -eq "Yes") {
                Write-Log "Setting ChildrenInheritToken for Policy '$Policy_Name'." "DEBUG"
                $PGApp.ChildrenInheritToken = $true
            }
            if ($RemoveAdminrightsfromFileOpen -eq "Yes") {
                Write-Log "Setting OpenDlgDropRights for Policy '$Policy_Name'." "DEBUG"
                $PGApp.OpenDlgDropRights = $true
            }

            # Add descriptions from CSV dictionary, if any
            if ($policyDictionary.ContainsKey($Policy_Name)) {
                Write-Log "Overriding description from dictionary for Policy '$Policy_Name'." "DEBUG"
                $PGApp.Description = $policyDictionary[$Policy_Name]
            }

            # Finally, add the application to the group
            $TargetAppGroup.Applications.Add($PGApp)
            $addedCount++
            $typeStats[$ApplicationType].Added++
            Write-Log "Successfully added application for Policy '$Policy_Name'." "INFO"

        }
        catch {
            $failedCount++

            # If we got as far as having an ApplicationType stat bucket, increment fail
            if ($ApplicationType -and $typeStats.ContainsKey($ApplicationType)) {
                $typeStats[$ApplicationType].Failed++
            }

            # Combine and shorten properties (except Policy Name) for a single-line log
            $allFields = @(
                "Policy Name: $Policy_Name"                                    # do NOT shorten
                "ApplicationType: $(Compress-String $ApplicationType)"
                "FileName: $(Compress-String $FileName)"
                "ChecksumAlg: $(Compress-String $ChecksumAlgorithm)"
                "Checksum: $(Compress-String $Checksum)"
                "SignedBy: $(Compress-String $SignedBy)"
                "Publisher: $(Compress-String $Publisher)"
                "ProductName: $(Compress-String $ProductName)"
                "FileDescription: $(Compress-String $FileDescription)"
                "ServiceName: $(Compress-String $ServiceName)"
                "WindowsAdminTask: $(Compress-String $WindowsAdminTask)"
                "Exception: $(Compress-String $_.Exception.Message)"
            ) -join " | "

            Write-Log "Line $lineCount - FAILED to add application. $allFields" "ERROR"
        }
    }

    # After the loop, if the last group was never added to PGConfig, add it
    if ($TargetAppGroup) {
        $PGConfig.ApplicationGroups.Add($TargetAppGroup)
        Write-Log "Added final AppGroup '$($TargetAppGroup.Name)' to PGConfig." "INFO"
    }
}
catch {
    Write-Log "Unexpected error during CSV processing: $($_.Exception.Message)" "ERROR"
}

### ---------------------------
### SAVE THE DEFENDPOINT POLICY
### ---------------------------

try {
    Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation $generatedXML
    Write-Log "Successfully saved Defendpoint settings to $(Split-Path $generatedXML -Leaf)."
}
catch {
    Write-Log "Failed to save $generatedXML : $($_.Exception.Message)" "ERROR"
}

### ---------------------------
### LOG SUMMARY
### ---------------------------

Write-Log "----- SUMMARY REPORT -----"
Write-Log "Total lines processed: $lineCount"
Write-Log "Successfully added applications: $addedCount"
Write-Log "Skipped applications (excluded patterns): $skippedCount"
Write-Log "Failed applications: $failedCount"
Write-Log "Not supported application types: $notSupportedCount"
Write-Log "Failed app groups (if any): $failedAppGroupCount"

if ($lineCount -gt 0) {
    $failPercentage = [Math]::Round(($failedCount / $lineCount * 100), 2)
    Write-Log "Failure percentage: $failPercentage %"
} else {
    Write-Log "No lines were processed."
}

Write-Log "----- DETAILED APPLICATION TYPE SUMMARY -----"
foreach ($appType in $typeStats.Keys) {
    $stats = $typeStats[$appType]
    Write-Log "$appType => Total: $($stats.Total), Added: $($stats.Added), Skipped: $($stats.Skipped), Failed: $($stats.Failed), NotSupported: $($stats.NotSupported)"
}

Write-Log "Script complete."
