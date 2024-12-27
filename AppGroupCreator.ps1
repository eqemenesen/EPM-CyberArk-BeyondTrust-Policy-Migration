### --------------------------
### SETUP & HELPER FUNCTIONS
### --------------------------

# If you'd like time-stamped logs in a consistent format, define a Write-Log function:
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

### ---------------------------
### IMPORT NECESSARY MODULES
### ---------------------------

try {
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -ErrorAction Stop
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -ErrorAction Stop
    Write-Log "Successfully imported Avecto modules."
}
catch {
    Write-Log "Failed to import Avecto modules. Error: $($_.Exception.Message)" "ERROR"
    return
}

### ---------------------------
### VARIABLES & INITIALIZATION
### ---------------------------

$reportFile = ".\PolicySummary_Garanti.csv"
$logFilePath = ".\logs\logfile_AppGroup.log"
$csvFilePath = ".\GarantiMainPolicy.csv"
$adminTasksFile = ".\AdminTasks.csv"  # example path for your admin tasks file

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

Write-Log "Retrieving Defendpoint Settings from blank_policy.xml"
try {
    $PGConfig = Get-DefendpointSettings -LocalFile -FileLocation "$baseFolder\blank_policy.xml"
    Write-Log "Loaded Defendpoint settings successfully."
}
catch {
    Write-Log "Failed to load blank_policy.xml: $($_.Exception.Message)" "ERROR"
    return
}

### ---------------------------
### IMPORT ADMIN TASKS
### ---------------------------
Write-Log "Importing admin tasks from $adminTasksFile"
$adminTasks = $null
try {
    $adminTasks = Import-Csv -Path $adminTasksFile -Delimiter ","
    Write-Log "Admin tasks loaded. Found $($adminTasks.Count) entries."
}
catch {
    Write-Log "Failed to load Admin Tasks CSV: $($_.Exception.Message)" "ERROR"
    return
}

### ---------------------------
### MAIN LOOP: CREATE APP GROUPS
### ---------------------------

# We will keep counters for advanced logging
[int]$lineCount   = 0
[int]$addedCount  = 0
[int]$skippedCount = 0
[int]$failedCount  = 0

Write-Log "Importing policy report from $reportFile"

$TargetAppGroupPre = "xxxxx"
$TargetAppGroup = $null
$processedAppGroups = @{}

try {
    Import-Csv -Path $reportFile -Delimiter "," | ForEach-Object {

        # Keep track of the line count
        $lineCount++

        # Grab all CSV properties
        $Policy_Name                      = $_."Policy Name"
        $ApplicationType                  = $_."Application Type"
        $FileName                         = $_."File Name"
        $FileNameCompareAs                = $_."File Name Compare As"
        $ChecksumAlgorithm                = $_."Checksum Algorithm"
        $Checksum                         = $_."Checksum"
        $Owner                            = $_."Owner"
        $ArgumentsCompareAs               = $_."Arguments Compare As"
        $SignedBy                         = $_."Signed By"
        $Publisher                        = $_."Publisher"
        $PublisherCompareAs               = $_."Publisher Compare As"
        $ProductName                      = $_."Product Name"
        $ProductNameCompareAs             = $_."Product Name Compare As"
        $FileDescription                  = $_."File Description"
        $FileDescriptionCompareAs         = $_."File Description Compare As"
        $CompanyName                      = $_."Company Name"
        $CompanyNameCompareAs             = $_."Company Name Compare As"
        $ProductVersionFrom               = $_."Product Version From"
        $ProductVersionTo                 = $_."Product Version To"
        $FileVersionFrom                  = $_."File Version From"
        $FileVersionTo                    = $_."File Version To"
        $ScriptFileName                   = $_."Script File Name"
        $ScriptShortcutName               = $_."Script Shortcut Name"
        $MSPProductName                   = $_."MSI/MSP Product Name"
        $MSPProductNameCompareAs          = $_."MSI/MSP Product Name Compare As"
        $MSPCompanyName                   = $_."MSI/MSP Company Name"
        $MSPCompanyNameCompareAs          = $_."MSI/MSP Company Name Compare As"
        $MSPProductVersionFrom            = $_."MSI/MSP Product Version From"
        $MSPProductVersionTo              = $_."MSI/MSP Product Version To"
        $MSPProductCode                   = $_."MSI/MSP Product Code"
        $MSPUpgradeCode                   = $_."MSI/MSP Upgrade Code"
        $WebApplicationURL                = $_."Web Application URL"
        $WebApplicationURLCompareAs       = $_."Web Application URL Compare As"
        $WindowsAdminTask                 = $_."Windows Admin Task"
        $AllActiveXInstallationAllowed    = $_."All ActiveX Installation Allowed"
        $ActiveXSourceURL                 = $_."ActiveX Source URL"
        $ActiveXSourceCompareAs           = $_."ActiveX Source Compare As"
        $ActiveXMIMEType                  = $_."ActiveX MIME Type"
        $ActiveXMIMETypeCompareAs         = $_."ActiveX MIME Type Compare As"
        $ActiveXCLSID                     = $_."ActiveX CLSID"
        $ActiveXVersionFrom               = $_."ActiveX Version From"
        $ActiveXVersionTo                 = $_."ActiveX Version To"
        $FileFolderPath                   = $_."File/Folder Path"
        $FileFolderType                   = $_."File/Folder Type"
        $FolderWithsubfoldersandfiles     = $_."Folder With subfolders and files"
        $RegistryKey                      = $_."Registry Key"
        $COMDisplayName                   = $_."COM Display Name"
        $COMCLSID                         = $_."COM CLSID"
        $ServiceName                      = $_."Service Name"
        $ElevatePrivilegesforChildProcesses = $_."Elevate Privileges for Child Processes"
        $RemoveAdminrightsfromFileOpen      = $_."Remove Admin rights from File Open/Save common dialogs"
        $ApplicationDescription             = $_."Application Description"
        
        Write-Host "Processing line: $lineCount, Policy Name: $Policy_Name"
        
        # We’ll wrap the creation of the PGApp object and property assignment in a try block
        try {
            # Skip known excluded lines
            if ($Policy_Name.Length -eq 0 -or 
                $Policy_Name -match "macOS" -or 
                $Policy_Name -match "Default MAC Policy" -or 
                $Policy_Name -clike 'Usage of "JIT*') {
                
                Write-Log "Line $lineCount - Skipped Policy '$Policy_Name' (excluded pattern)." "WARN"
                $skippedCount++
                return
            }

            # Create the new application object
            $PGApp = New-Object Avecto.Defendpoint.Settings.Application $PGConfig

            # Check and prevent duplicate app group creation
            if (-not $processedAppGroups.ContainsKey($Policy_Name)) {
                if ($Policy_Name -ne $TargetAppGroupPre) {
                    # If there is an existing group, add it to the PGConfig
                    if ($TargetAppGroup) {
                        $PGConfig.ApplicationGroups.Add($TargetAppGroup)
                        Write-Log "Added AppGroup '$($TargetAppGroup.Name)' to PGConfig."
                    }

                    # Create a new group
                    $TargetAppGroup = New-Object Avecto.Defendpoint.Settings.ApplicationGroup
                    $TargetAppGroup.Name        = $Policy_Name
                    $TargetAppGroup.Description = $Policy_Name
                    $TargetAppGroupPre          = $Policy_Name

                    # Mark this app group as processed
                    $processedAppGroups[$Policy_Name] = $true
                }
            }

            # If we have a custom description, use it
            if ($ApplicationDescription.Length -ge 1) {
                $PGApp.Description = $ApplicationDescription
            } elseif ($FileName.Length -ge 1) {
                $PGApp.Description = $FileName
            }

            switch -Wildcard ($ApplicationType) {
                "ActiveX Control" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ActiveX
                }
                "Admin Tasks" {
                    $AdminApplicationName = $WindowsAdminTask
                    $adminTask = $adminTasks | Where-Object { $_.AdminApplicationName -eq $AdminApplicationName }

                    if ("" -eq $adminTask) {
                        Write-Log "Line $lineCount - Admin Task '$AdminApplicationName' not found in adminTasksFile." "WARN"
                        $failedCount++
                        return
                    }

                    switch ($adminTask.AdminType) {
                        "Executable" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Executable
                            $PGApp.AppxPackageNameMatchCase = "Contains"
                            $PGApp.FileName  = $adminTask.AdminPath
                            $PGApp.ProductName = $adminTask.AdminApplicationName
                            $PGApp.DisplayName = $adminTask.AdminApplicationName
                            if ($adminTask.AdminCommandLine -ne $null) {
                                $PGApp.CmdLine = $adminTask.AdminCommandLine
                                $PGApp.CmdLineMatchCase = "false"
                                $PGApp.CmdLineStringMatchType = "Contains"
                            }
                        }
                        "COMClass" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::COMClass
                            $PGApp.DisplayName = $adminTask.AdminApplicationName
                            $PGApp.CheckAppID = "true"
                            $PGApp.CheckCLSID = "true"
                            $PGApp.AppID = $adminTask.AdminCLSID
                            $PGApp.CLSID = $adminTask.AdminCLSID
                        }
                        "ManagementConsoleSnapin" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ManagementConsoleSnapin
                            $PGApp.DisplayName = $adminTask.AdminApplicationName
                            $PGApp.AppxPackageNameMatchCase = "Contains"
                            $PGApp.FileName = $adminTask.AdminPath
                        }
                        "ControlPanelApplet" {
                            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ControlPanelApplet
                            $PGApp.CheckAppID = "true"
                            $PGApp.CheckCLSID = "true"
                            $PGApp.AppID = $adminTask.AdminCLSID
                            $PGApp.CLSID = $adminTask.AdminCLSID
                        }
                        default {
                            Write-Log "Line $lineCount - Policy '$Policy_Name', AdminType '$($adminTask.AdminType)' is not supported." "WARN"
                            $failedCount++
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
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Dll
                    Write-Log "Line $lineCount - Policy '$Policy_Name', 'Dynamic-Link Library' not supported." "WARN"
                    $skippedCount++
                    return

                }
                "Service" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Service
                }
                "Registry Key" {
                    $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::RegistrySettings
                }
                "File or Directory System Entry" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is NOT supported (error)." "ERROR"
                    $failedCount++
                    return
                }
                "Microsoft Update (MSU)" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is NOT supported (error)." "ERROR"
                    $failedCount++
                    return
                }
                "Web Application" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is NOT supported (error)." "ERROR"
                    $failedCount++
                    return
                }
                "Win App" {
                    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is NOT supported (error)." "ERROR"
                    $failedCount++
                    return
                }
                #Default {
                #    Write-Log "Line $lineCount - Policy '$Policy_Name', '$ApplicationType' is UNRECOGNIZED (error)." "ERROR"
                #    $failedCount++
                #    return
                #}
            }

            # File Name
            if ($FileName) {
                $PGApp.CheckFileName = 1
                if ($FileName.StartsWith("*")) {
                    # EndsWith
                    $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                    $PGApp.FileName = $FileName.Substring(1)
                }
                elseif ($FileName.EndsWith("*")) {
                    # StartsWith
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
                $PGApp.CheckFileHash = 1
                $PGApp.FileHash = $Checksum
            }

            # Publisher
            if ($SignedBy -eq "Specific Publishers" -and $Publisher) {
                $PGApp.CheckPublisher = 1
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
                $PGApp.CheckProductName = 1
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
                $PGApp.CheckProductDesc = 1
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
                $PGApp.CheckMinProductVersion = 1
                $PGApp.MinProductVersion = $ProductVersionFrom
            }
            if ($ProductVersionTo) {
                $PGApp.CheckMaxProductVersion = 1
                $PGApp.MaxProductVersion = $ProductVersionTo
            }

            # File Version From/To
            if ($FileVersionFrom) {
                $PGApp.CheckMinFileVersion = 1
                $PGApp.MinFileVersion = $FileVersionFrom
            }
            if ($FileVersionTo) {
                $PGApp.CheckMaxFileVersion = 1
                $PGApp.MaxFileVersion = $FileVersionTo
            }

            # Service
            if ($ServiceName) {
                $PGApp.CheckServiceName = 1
                $PGApp.ServiceNamePatternMatching = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                $PGApp.ServiceName = $ServiceName
                $PGApp.ServicePause = 1
                $PGApp.ServiceStart = 1
                $PGApp.ServiceStop = 1
                $PGApp.ServiceConfigure = 1
            }

            # Elevate / Remove Admin rights
            if ($ElevatePrivilegesforChildProcesses -eq "Yes") {
                $PGApp.ChildrenInheritToken = 1
            }
            if ($RemoveAdminrightsfromFileOpen -eq "Yes") {
                $PGApp.OpenDlgDropRights = 1
            }

            # Add descriptions from CSV dictionary, if any
            if ($policyDictionary.ContainsKey($Policy_Name)) {
                $PGApp.Description = $policyDictionary[$Policy_Name]
            }

            # Add the final application to the group
            $TargetAppGroup.Applications.Add($PGApp)
            $addedCount++
        }
        catch {
            # If something fails in creating or adding the application, log it
            $failedCount++

            Write-Log "Line $lineCount - FAILED to add application with Policy '$Policy_Name'. See details below." "ERROR"
            
            # Let's log all relevant properties for debugging
            Write-Log " --- Policy_Name: $Policy_Name" "DEBUG"
            Write-Log " --- ApplicationType: $ApplicationType" "DEBUG"
            Write-Log " --- FileName: $FileName" "DEBUG"
            Write-Log " --- ChecksumAlgorithm: $ChecksumAlgorithm" "DEBUG"
            Write-Log " --- Checksum: $Checksum" "DEBUG"
            Write-Log " --- SignedBy: $SignedBy" "DEBUG"
            Write-Log " --- Publisher: $Publisher" "DEBUG"
            Write-Log " --- ProductName: $ProductName" "DEBUG"
            Write-Log " --- FileDescription: $FileDescription" "DEBUG"
            Write-Log " --- ProductVersionFrom: $ProductVersionFrom" "DEBUG"
            Write-Log " --- ProductVersionTo: $ProductVersionTo" "DEBUG"
            Write-Log " --- FileVersionFrom: $FileVersionFrom" "DEBUG"
            Write-Log " --- FileVersionTo: $FileVersionTo" "DEBUG"
            Write-Log " --- WindowsAdminTask: $WindowsAdminTask" "DEBUG"
            Write-Log " --- ServiceName: $ServiceName" "DEBUG"
            Write-Log " --- ElevatePrivilegesForChildProcesses: $ElevatePrivilegesforChildProcesses" "DEBUG"
            Write-Log " --- RemoveAdminrightsfromFileOpen: $RemoveAdminrightsfromFileOpen" "DEBUG"
            Write-Log " --- Exception: $($_.Exception.Message)" "ERROR"
        }
    }

    # After the loop, if the last group is created but not added yet
    if ($TargetAppGroup) {
        $PGConfig.ApplicationGroups.Add($TargetAppGroup)
        Write-Log "Added final AppGroup '$($TargetAppGroup.Name)' to PGConfig."
    }
}
catch {
    Write-Log "Unexpected error during CSV processing: $($_.Exception.Message)" "ERROR"
}

### ---------------------------
### SAVE THE DEFENDPOINT POLICY
### ---------------------------

try {
    Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation "$baseFolder\generated_appGroup.xml"
    Write-Log "Successfully saved Defendpoint settings to generated_appGroup.xml."
}
catch {
    Write-Log "Failed to save generated_appGroup.xml: $($_.Exception.Message)" "ERROR"
}

### ---------------------------
### LOG SUMMARY
### ---------------------------

Write-Log "----- SUMMARY REPORT -----"
Write-Log "Total lines processed: $lineCount"
Write-Log "Successfully added applications: $addedCount"
Write-Log "Skipped applications: $skippedCount"
Write-Log "Failed applications: $failedCount"

if ($lineCount -gt 0) {
    $failPercentage = [Math]::Round(($failedCount / $lineCount * 100), 2)
    Write-Log "Failure percentage: $failPercentage %"
} else {
    Write-Log "No lines were processed."
}

Write-Log "Script complete."
