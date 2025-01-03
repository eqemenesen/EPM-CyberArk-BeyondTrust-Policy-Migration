param(
    [string]$baseFolder = ".\Areas\GarantiTest"
)

### -----------------------------------------
### IMPORT REQUIRED MODULES
### -----------------------------------------
try {
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -Force -ErrorAction Stop
    Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -Force -ErrorAction Stop
    Write-Host "Successfully imported Avecto modules."
} catch {
    Write-Host "ERROR: Failed to import Avecto modules. $_"
    exit 1
}

### -----------------------------------------
### HELPER: LOG & STRING SHORTENER FUNCTIONS
### -----------------------------------------

# Log function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry  = "$timestamp [$Level] $Message"
    $logEntry  | Out-File -FilePath $logFilePath -Append
    Write-Host $logEntry
}

# Shorten a string to a given length. Useful if you have properties that can be very long.
function ShortenString {
    param(
        [string]$inputString,
        [int]$MaxLength = 30
    )

    if ([string]::IsNullOrEmpty($inputString)) {
        return $inputString
    }
    if ($inputString.Length -le $MaxLength) {
        return $inputString
    }
    else {
        return $inputString.Substring(0, $MaxLength) + "..."
    }
}

### -----------------------------------------
### MAIN SCRIPT
### -----------------------------------------

# Paths
$logFilePath     = "$baseFolder\logs\logFile_ContenGroup.log"
$outputFilePath  = "$baseFolder\generated_AppAndContentGroups.xml"
$blankPolicyFile = "$baseFolder\generated_appGroup.xml"
$reportFile      = "$baseFolder\PolicySummary_Garanti.csv"
$csvFilePath     = "$baseFolder\GarantiMainPolicy.csv"

# Ensure the log directory exists
if (-not (Test-Path -Path (Split-Path -Path $logFilePath))) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Path $logFilePath)
}

# Clear old log
"" | Out-File -FilePath $logFilePath -Force

Write-Log "----- Script started for Content Groups -----" "INFO"

### -----------------------------------------
### BUILD A POLICY DICTIONARY
### -----------------------------------------
Write-Log "Importing ticket numbers from $csvFilePath" "INFO"

$policyDictionary = @{}

try {
    Import-Csv -Path $csvFilePath | ForEach-Object {
        $policyName        = $_."Policy Name"
        $policyDescription = $_."Policy Description"
        
        if ($policyName) {
            # Clean multiline descriptions
            $cleanedDescription = $policyDescription -replace "`r`n|`n|`r", " - "
            if ($cleanedDescription) {
                $policyDictionary[$policyName] = $cleanedDescription
            }
        }
    }
    Write-Log "Successfully built policy dictionary from $csvFilePath." "INFO"
}
catch {
    Write-Log "Error building policy dictionary. Error: $($_.Exception.Message)" "ERROR"
    return
}

### -----------------------------------------
### LOAD EXISTING POLICY CONFIG
### -----------------------------------------
try {
    $PGConfig = Get-DefendpointSettings -LocalFile -FileLocation $blankPolicyFile
    Write-Log "Successfully loaded existing policy configuration: $blankPolicyFile" "INFO"
}
catch {
    Write-Log "Failed to load policy file '$blankPolicyFile'. Error: $($_.Exception.Message)" "ERROR"
    return
}

### -----------------------------------------
### INITIALIZE VARIABLES & COUNTERS
### -----------------------------------------
$processedContentGroups = @{}
[int]$lineCount       = 0
[int]$addedCount      = 0  # number of content entries successfully added
[int]$skippedCount    = 0  # number of content rows skipped
[int]$failedCount     = 0  # number of content entries that failed creation
[int]$notContentCount = 0  # number of rows that are not "File or directory System Entry"

Write-Log "Reading CSV report from $reportFile" "INFO"

### -----------------------------------------
### PROCESS THE CSV
### -----------------------------------------
try {
    $csvData = Import-Csv -Path $reportFile

    foreach ($row in $csvData) {
        $lineCount++

        # Gather properties
        $policyName        = $row."Policy Name"
        $applicationType   = $row."Application Type"
        $fileFolderPath    = $row."File/Folder Path"
        $fileFolderType    = $row."File/Folder Type"

        Write-Log "Line $lineCount => Policy Name: $policyName, AppType: $applicationType, File/Folder Path: $fileFolderPath, File/Folder Type: $fileFolderType" "DEBUG"

        # We only care about rows where "Application Type" == "File or directory System Entry"
        if ($applicationType -ne "File or directory System Entry") {
            Write-Log "Line $lineCount => Skipped (AppType $applicationType != 'File or directory System Entry')" "DEBUG"
            $notContentCount++
            continue
        }

        # Create or reuse a ContentGroup for this policy
        try {
            if (-not $processedContentGroups.ContainsKey($policyName)) {
                # New group
                Write-Log "Line $lineCount => Creating new ContentGroup for Policy Name '$policyName'." "INFO"
                $contentGroup = New-Object Avecto.Defendpoint.Settings.ContentGroup
                $contentGroup.Name = $policyName

                # If there's a custom description in the dictionary, use it
                if ($policyDictionary.ContainsKey($policyName)) {
                    $contentGroup.Description = $policyDictionary[$policyName]
                }
                else {
                    $contentGroup.Description = "Content group for policy: $policyName"
                }

                $PGConfig.ContentGroups.Add($contentGroup)
                $processedContentGroups[$policyName] = $contentGroup
            }
            else {
                # Reuse existing group
                Write-Log "Line $lineCount => Using existing ContentGroup for Policy Name '$policyName'." "DEBUG"
                $contentGroup = $processedContentGroups[$policyName]
            }
        }
        catch {
            Write-Log "Line $lineCount => Failed to create or retrieve ContentGroup for '$policyName'. Error: $($_.Exception.Message)" "ERROR"
            $failedCount++
            continue
        }

        # Now, create a new Content object for the file/folder info
        try {
            # Avecto requires an internal config object
            $config  = [Avecto.Defendpoint.Settings.Configuration]::new()
            $content = [Avecto.Defendpoint.Settings.Content]::new($config)

            $content.ID          = [guid]::NewGuid()
            $content.Description = $policyName
            if ($policyDictionary.ContainsKey($policyName)) {
                $content.Description = $policyDictionary[$policyName]
            }

            # We'll log the path for debugging
            Write-Log "Line $lineCount => Setting FileName = '$fileFolderPath' for Content object." "DEBUG"
            $content.FileName         = $fileFolderPath
            $content.CheckFileName    = $true

            # Determine matching type
            switch ($fileFolderType) {
                "dir" {
                    Write-Log "Line $lineCount => FileFolderType='dir', setting FileNameStringMatchType=StartsWith." "DEBUG"
                    $content.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::StartsWith
                }
                "file" {
                    Write-Log "Line $lineCount => FileFolderType='file', setting FileNameStringMatchType=Exact." "DEBUG"
                    $content.FileNameStringMatchType = [Avecto.Defendpoint.Settings.StringMatchType]::Exact
                }
                default {
                    Write-Log "Line $lineCount => Unrecognized FileFolderType '$fileFolderType'. Skipping creation." "WARN"
                    $skippedCount++
                    continue
                }
            }

            # Finally add Content to the group
            $contentGroup.Contents.Add($content)
            Write-Log "Line $lineCount => Successfully added new Content object (ID=$($content.ID)) to ContentGroup '$policyName'." "INFO"
            $addedCount++
        }
        catch {
            $failedCount++
            # We'll log all relevant properties in one line, using ShortenString if necessary
            $summarizedRow = @(
                "PolicyName: $policyName"   # do not shorten policyName
                "AppType: $(ShortenString $applicationType)"
                "Path: $(ShortenString $fileFolderPath)"
                "Type: $(ShortenString $fileFolderType)"
                "Error: $(ShortenString $_.Exception.Message)"
            ) -join " | "

            Write-Log "Line $lineCount => FAILED to add content. $summarizedRow" "ERROR"
        }
    }
}
catch {
    Write-Log "Line $lineCount => Error processing CSV file '$reportFile'. Exception: $($_.Exception.Message)" "ERROR"
    return
}

### -----------------------------------------
### SAVE UPDATED POLICY
### -----------------------------------------
try {
    Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation $outputFilePath
    Write-Log "Successfully saved updated policy to '$outputFilePath'." "INFO"
}
catch {
    Write-Log "Failed to save updated policy file '$outputFilePath'. Error: $($_.Exception.Message)" "ERROR"
}

### -----------------------------------------
### LOG SUMMARY
### -----------------------------------------
Write-Log "----- SUMMARY REPORT -----" "INFO"
Write-Log "Total lines processed: $lineCount" "INFO"
Write-Log "Content items added:   $addedCount" "INFO"
Write-Log "Content items skipped: $skippedCount" "INFO"
Write-Log "Content items failed:  $failedCount" "INFO"
Write-Log "Non-content rows:      $notContentCount" "INFO"

if ($lineCount -gt 0) {
    $failPercentage = [Math]::Round(($failedCount / $lineCount * 100), 2)
    Write-Log "Failure percentage:    $failPercentage %" "INFO"
} else {
    Write-Log "No lines were processed." "WARN"
}

Write-Log "Script completed." "INFO"
