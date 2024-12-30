# Define paths
$logFilePath = ".\logs\logFile_ContenGroup.log"
$outputFilePath = ".\generated_AppAndContentGroups.xml"
$blankPolicyFile = ".\generated_appGroup.xml"
$reportFile = ".\PolicySummary_Garanti.csv"
$csvFilePath = ".\GarantiMainPolicy.csv"

# Ensure the log directory exists
if (-not (Test-Path -Path (Split-Path -Path $logFilePath))) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Path $logFilePath)
}


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


# Clear the old log file
"" | Out-File -FilePath $logFilePath -Force

# Log function
function Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp [$Level] $Message"
    $logEntry | Out-File -FilePath $logFilePath -Append
    Write-Host $logEntry
}

# Start script logging
Log "Script started."

# Import required modules
try {
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -ErrorAction Stop
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -ErrorAction Stop
    Write-Log "Successfully imported Avecto modules."
}
catch {
    Write-Log "Failed to import Avecto modules. Error: $($_.Exception.Message)" "ERROR"
    return
}

# Load existing policy configuration
try {
    $PGConfig = Get-DefendpointSettings -LocalFile -FileLocation $blankPolicyFile
    Log "Successfully loaded existing policy configuration."
} catch {
    Log "Failed to load policy file: $blankPolicyFile. Error: $($_.Exception.Message)" "ERROR"
    return
}

# Initialize processed ContentGroups
$processedContentGroups = @{}
$lineCount = 0

# Read and process the CSV file
try {
    $csvData = Import-Csv -Path $reportFile
    foreach ($row in $csvData) {
        $policyName = $row."Policy Name"
        $applicationType = $row."Application Type"
        $fileFolderPath = $row."File/Folder Path"
        $fileFolderType = $row."File/Folder Type"

        # Process only rows with "Application Type" equal to "File or directory System Entry"
        if ($applicationType -ne "File or directory System Entry") {
            Log "$lineCount Skipped processing row for Policy Name '$policyName' as Application Type is '$applicationType'." "DEBUG"
            $lineCount++
            continue
        }

        # Create or reuse a ContentGroup for the Policy Name
        if (-not $processedContentGroups.ContainsKey($policyName)) {
            $contentGroup = New-Object Avecto.Defendpoint.Settings.ContentGroup
            $contentGroup.Name = $policyName
            $contentGroup.Description = "Content group for policy: $policyName"

            if ($policyDictionary.ContainsKey($policyName)) {
                $contentGroup.Description = $policyDictionary[$policyName]
            }
            
            $PGConfig.ContentGroups.Add($contentGroup)
            $processedContentGroups[$policyName] = $contentGroup
            Log "$lineCount Created new ContentGroup for Policy Name '$policyName'."
        } else {
            $contentGroup = $processedContentGroups[$policyName]
            Log "$lineCount Reusing existing ContentGroup for Policy Name '$policyName'."
        }

        # Create a new Content
        $config = [Avecto.Defendpoint.Settings.Configuration]::new()
        $content = [Avecto.Defendpoint.Settings.Content]::new($config)
        $content.ID = [guid]::NewGuid()
        $content.Description = $policyName

        if ($policyDictionary.ContainsKey($policyName)) {
            $content.Description = $policyDictionary[$policyName]
        }

        $content.FileName = $fileFolderPath
        $content.CheckFileName = $true

        # Set FileNameStringMatchType based on File/Folder Type
        if ($fileFolderType -eq "dir") {
            $content.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
        } elseif ($fileFolderType -eq "file") {
            $content.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
        } else {
            Log "$lineCount Skipped content creation for Policy Name '$policyName' due to invalid File/Folder Type: $fileFolderType." "WARN"
            $lineCount++
            continue
        }

        # Add Content to the ContentGroup
        $contentGroup.Contents.Add($content)
        Log "$lineCount Added Content to ContentGroup '$policyName' with FileNameStringMatchType '$($content.FileNameStringMatchType)'."
        $lineCount++
    }
} catch {
    Log "$lineCount Error processing CSV file: $($_.Exception.Message)" "ERROR"
    return
}

# Save the updated policy configuration
try {
    Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation $outputFilePath
    Log "Successfully saved updated policy to $outputFilePath."
} catch {
    Log "Failed to save updated policy file: $outputFilePath. Error: $($_.Exception.Message)" "ERROR"
}

Log "Script completed."
