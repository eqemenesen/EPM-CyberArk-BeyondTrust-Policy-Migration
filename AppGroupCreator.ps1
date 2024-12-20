Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll'
Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll'
$baseFolder = "."
$reportFile = "$baseFolder\PolicySummary_Garanti.csv"
$logFilePath = "$baseFolder\logfile_AppGroup.txt"

Write-Host $logFilePath , $reportFile
"" | Out-File -FilePath $logFilePath -Force
#$csv = Import-Csv -path $path -Header "Original File Name","Checksum"





# Ticket nnumaralarının almak için import
# Define the CSV file path
$csvFilePath = "./GarantiMainPolicy.csv"

# Initialize an empty dictionary (hashtable)
$policyDictionary = @{}
Import-Csv -Path $csvFilePath | ForEach-Object {
    # Ensure "Policy Name" and "Policy Description" exist and add them to the dictionary
    $policyName = $_."Policy Name"
    $policyDescription = $_."Policy Description"
    
    if ($policyName) {
        # Replace newlines in Policy Description with " - "
        $cleanedDescription = $policyDescription -replace "`r`n|`n|`r", " - "
        
        # Add to the dictionary only if description is non-empty
        if ($cleanedDescription) {
            $policyDictionary[$policyName] = $cleanedDescription
        }
    }
}




# Get settings
$PGConfig = Get-DefendpointSettings -LocalFile -FileLocation "$baseFolder\blank_policy.xml"

$adminTasks = Import-Csv -Path $adminTasksFile -Delimiter ","

# Find target Application Group
$line = 0
$TargetAppGroupPre = "xxxxx"
$TargetAppGroup = ''
Import-Csv -path $reportFile -Delimiter "," | Foreach-Object {
        
        # Get all columns from policy report

        $Policy_Name = $_."Policy Name"
        $ApplicationType = $_."Application Type"
        $FileName = $_."File Name"
        $FileNameCompareAs = $_."File Name Compare As"
        $ChecksumAlgorithm = $_."Checksum Algorithm"
        $Checksum = $_."Checksum"
        $Owner = $_."Owner"
        $ArgumentsCompareAs = $_."Arguments Compare As"
        $SignedBy = $_."Signed By"
        $Publisher = $_."Publisher"
        $PublisherCompareAs = $_."Publisher Compare As"
        $ProductName = $_."Product Name"
        $ProductNameCompareAs = $_."Product Name Compare As"
        $FileDescription = $_."File Description"
        $FileDescriptionCompareAs = $_."File Description Compare As"
        $CompanyName = $_."Company Name"
        $CompanyNameCompareAs = $_."Company Name Compare As"
        $ProductVersionFrom = $_."Product Version From"
        $ProductVersionTo = $_."Product Version To"
        $FileVersionFrom = $_."File Version From"
        $FileVersionTo = $_."File Version To"
        $ScriptFileName = $_."Script File Name"
        $ScriptShortcutName = $_."Script Shortcut Name"
        $MSPProductName = $_."MSI/MSP Product Name"
        $MSPProductNameCompareAs = $_."MSI/MSP Product Name Compare As"
        $MSPCompanyName = $_."MSI/MSP Company Name"
        $MSPCompanyNameCompareAs = $_."MSI/MSP Company Name Compare As"
        $MSPProductVersionFrom = $_."MSI/MSP Product Version From"
        $MSPProductVersionTo = $_."MSI/MSP Product Version To"
        $MSPProductCode = $_."MSI/MSP Product Code"
        $MSPUpgradeCode = $_."MSI/MSP Upgrade Code"
        $WebApplicationURL = $_."Web Application URL"
        $WebApplicationURLCompareAs = $_."Web Application URL Compare As"
        $WindowsAdminTask = $_."Windows Admin Task"
        $AllActiveXInstallationAllowed = $_."All ActiveX Installation Allowed"
        $ActiveXSourceURL = $_."ActiveX Source URL"
        $ActiveXSourceCompareAs = $_."ActiveX Source Compare As"
        $ActiveXMIMEType = $_."ActiveX MIME Type"
        $ActiveXMIMETypeCompareAs = $_."ActiveX MIME Type Compare As"
        $ActiveXCLSID = $_."ActiveX CLSID"
        $ActiveXVersionFrom = $_."ActiveX Version From"
        $ActiveXVersionTo = $_."ActiveX Version To"
        $FileFolderPath = $_."File/Folder Path"
        $FileFolderType = $_."File/Folder Type"
        $FolderWithsubfoldersandfiles = $_."Folder With subfolders and files"
        $RegistryKey = $_."Registry Key"
        $COMDisplayName = $_."COM Display Name"
        $COMCLSID = $_."COM CLSID"
        $ServiceName = $_."Service Name"
        $ElevatePrivilegesforChildProcesses = $_."Elevate Privileges for Child Processes"
        $RemoveAdminrightsfromFileOpen = $_."Remove Admin rights from File Open/Save common dialogs"
        $ApplicationDescription = $_."Application Description"

        if ($Policy_Name.Length -gt 0 -and -not ($Policy_Name -match "macOS" -or $Policy_Name -match "Default MAC Policy")) {
            $PGApp = New-Object Avecto.Defendpoint.Settings.Application $PGConfig
        } else {
            return
        }        
        
        Write-Host $line, $Policy_Name
        
        if ($Policy_Name -ne $TargetAppGroupPre){
            $PGConfig.ApplicationGroups.Add($TargetAppGroup)
            $TargetAppGroup = new-object Avecto.Defendpoint.Settings.ApplicationGroup
            $TargetAppGroup.Name = $Policy_Name
            $TargetAppGroup.Description = $Policy_Name
            $TargetAppGroupPre = $Policy_Name
        }
		
        if ($ApplicationDescription.Length -ge 1){
            $PGApp.Description = $ApplicationDescription
        }elseif($FileName.Length -ge 1){
            $PGApp.Description = $FileName
        }

        if ( $ApplicationType -eq "ActiveX Control" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ActiveX
            $TargetAppGroup.Description = $Policy_Name
            $TargetAppGroupPre = $Policy_NameControl
        }elseif ( $ApplicationType -eq "Admin Tasks" )
        {
            $AdminApplicationName = $WindowsAdminTask

            $adminTask = $adminTasks | Where-Object { $_.AdminApplicationName -eq $AdminApplicationName }

            if ($null -eq $adminTask) {
                $logEntry = "line: $line, Admin Task $AdminApplicationName not supported." | Add-Content -Path $logFilePath
                return
            }

            # Add the admin task to the group
            if ($adminTask.AdminType -eq "Executable") {
                $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Executable
                $PGApp.AppxPackageNameMatchCase = "Contains"
                $PGApp.FileName = $adminTask.AdminPath
                $PGApp.ProductName = $adminTask.AdminApplicationName
                $PGApp.DisplayName = $adminTask.AdminApplicationName
                if ($adminTask.AdminCommandLine -ne $null){
                    $PGApp.CmdLine = $adminTask.AdminCommandLine
                    $PGApp.CmdLineMatchCase = "false"
                    $PGApp.CmdLineStringMatchType = "Contains"
                }
                
            } elseif ($adminTask.AdminType -eq "COMClass") {
                $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::COMClass
                $PGApp.DisplayName = $adminTask.AdminApplicationName
                $PGApp.CheckAppID = "true"
                $PGApp.CheckCLSID = "true"
                $PGApp.AppID = $adminTask.AdminCLSID
                $PGApp.CLSID = $adminTask.AdminCLSID

            } elseif ($adminTask.AdminType -eq "ManagementConsoleSnapin") {
                $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ManagementConsoleSnapin
                $PGApp.DisplayName = $adminTask.AdminApplicationName
                $PGApp.AppxPackageNameMatchCase = "Contains"
                $PGApp.FileName = $adminTask.AdminPath

                
            } elseif ($adminTask.AdminType -eq "ControlPanelApplet") {
                $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::ControlPanelApplet
                $PGApp.CheckAppID = "true"
                $PGApp.CheckCLSID = "true"
                $PGApp.AppID = $adminTask.AdminCLSID
                $PGApp.CLSID = $adminTask.AdminCLSID
            }
            
        }elseif ( $ApplicationType -eq "Application Group" )
        {
            $logEntry = "line: $line, Policy Name : $Policy_Name, $ApplicationType not supported." | Add-Content -Path $logFilePath
            return
        }elseif ( $ApplicationType -eq "Executable" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Executable
        }elseif ( $ApplicationType -eq "COM Object" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::COMClass
        }elseif ( $ApplicationType -eq "MSI/MSP Installation" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::InstallerPackage
        }elseif ( $ApplicationType -eq "Script" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::BatchFile
        }elseif ( $ApplicationType -eq "Dynamic-Link Library" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Dll
        }elseif ( $ApplicationType -eq "Service" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::Service
        } elseif ( $ApplicationType -eq "File or Directory System Entry" )
        {
            $logEntry = "line: $line, Policy Name : $Policy_Name, $ApplicationType not supported." | Add-Content -Path $logFilePath
            return
        }elseif ( $ApplicationType -eq "Microsoft Update (MSU)" )
        {
            $logEntry = "line: $line, Policy Name : $Policy_Name, $ApplicationType not supported." | Add-Content -Path $logFilePath
            return
        }elseif ( $ApplicationType -eq "Registry Key" )
        {
            $PGApp.Type = [Avecto.Defendpoint.Settings.ApplicationType]::RegistrySettings
        }elseif ( $ApplicationType -eq "Web Application" )
        {
            $logEntry = "line: $line, Policy Name : $Policy_Name, $ApplicationType not supported." | Add-Content -Path $logFilePath
            return
        }elseif ( $ApplicationType -eq "Win App" )
        {
            $logEntry = "line: $line, Policy Name : $Policy_Name, $ApplicationType not supported." | Add-Content -Path $logFilePath
            return
        }

        #File Name
        if ($FileName.Length -gt 0){
            $PGApp.CheckFileName = 1   
            if ($FileName[0] -eq "*")
            {
                $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                $PGApp.FileName = $FileName.Substring(1)
            }
            elseif ($FileName[$FileName.Length - 1] -eq "*")
            {
                $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                $PGApp.FileName = $FileName.Substring(0, $FileName.Length - 1)
            }else{
                $PGApp.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                $PGApp.FileName = $FileName
            }
            
        }
        
        # CheckSum 
        if ($ChecksumAlgorithm -eq "SHA1"){
            $PGApp.CheckFileHash = 1
            $PGApp.FileHash = $Checksum
        }

        # Publisher
        if ($SignedBy -eq "Specific Publishers"){
            $PGApp.CheckPublisher = 1 
            if ($Publisher[0] -eq "*")
            {
                $PGApp.PublisherStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                $PGApp.Publisher = $Publisher.Substring(1)
            }
            elseif ($Publisher[$Publisher.Length - 1] -eq "*")
            {
                $PGApp.PublisherStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                $PGApp.Publisher = $Publisher.Substring(0, $Publisher.Length - 1)
            }else{
                $PGApp.PublisherStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                $PGApp.Publisher = $Publisher
            }

        }

        # Product Name
        if ($ProductName.Length -gt 0){
            $PGApp.CheckProductName = 1 
            if ($ProductName[0] -eq "*")
            {
                $PGApp.ProductNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                $PGApp.ProductName = $ProductName.Substring(1)
            }
            elseif ($ProductName[$ProductName.Length - 1] -eq "*")
            {
                $PGApp.ProductNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                $PGApp.ProductName = $ProductName.Substring(0, $ProductName.Length - 1)
            }else{
                $PGApp.ProductNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                $PGApp.ProductName = $ProductName
            }

        }

        
        # Product Description
        if ($FileDescription.Length -gt 0){
            $PGApp.ProductDesc = 1 
            if ($FileDescription[0] -eq "*")
            {
                $PGApp.ProductDescStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::EndsWith
                $PGApp.ProductDesc = $FileDescription.Substring(1)
            }
            elseif ($FileDescription[$FileDescription.Length - 1] -eq "*")
            {
                $PGApp.ProductDescStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                $PGApp.ProductDesc = $FileDescription.Substring(0, $FileDescription.Length - 1)
            }else{
                $PGApp.ProductDescStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                $PGApp.ProductDesc = $FileDescription
            }

        }


        # Product Version From
        if ($ProductVersionFrom.Length -gt 0){
            $PGApp.CheckMinProductVersion = 1 
            $PGApp.MinProductVersion = $ProductVersionFrom
        }
        # Product Version To
        if ($ProductVersionTo.Length -gt 0){
            $PGApp.CheckMaxProductVersion = 1 
            $PGApp.MaxProductVersion = $ProductVersionTo
        }

        # File Version From
        if ($FileVersionFrom.Length -gt 0){
            $PGApp.CheckMinFileVersion = 1 
            $PGApp.MinFileVersion = $FileVersionFrom
        }
        # File Version To
        if ($ProductVersionTo.Length -gt 0){
            $PGApp.CheckMaxFileVersion = 1 
            $PGApp.MaxFileVersion = $FileVersionFrom
        }

        #Service
        if ($ServiceName.Length -gt 0){
            $PGApp.CheckServiceName = 1 
            $PGApp.ServiceNamePatternMatching = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
            $PGApp.ServiceName = $ServiceName
            $PGApp.ServicePause = 1
            $PGApp.ServiceStart = 1 
            $PGApp.ServiceStop = 1 
            $PGApp.ServiceConfigure = 1
        }

        If ($ElevatePrivilegesforChildProcesses -eq "Yes") {
                $PGApp.ChildrenInheritToken=1
            }
        If ($RemoveAdminrightsfromFileOpen -eq "Yes") {
                $PGApp.OpenDlgDropRights=1
            }
        
        # THIS PART IS ONLY FOR ADDING DESSCRIPTIONS
        if ($policyDictionary.ContainsKey($Policy_Name)) {
            $PGApp.Description = $policyDictionary[$Policy_Name]
        }

        $TargetAppGroup.Applications.Add($PGApp)
        $line++
        
        
}
Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation "$baseFolder\generated_appGroup.xml"