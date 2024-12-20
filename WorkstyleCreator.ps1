# Import necessary modules
Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -Force
Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -Force

# Define paths
$baseFolder = "."
$reportFile = "$baseFolder\GarantiMainPolicy.csv"
$logFilePath = "$baseFolder\logfile_Workstyle.txt"
$outputFile = "$baseFolder/generated_policy_with_assignments.xml"

Write-Host $logFilePath , $reportFile
"" | Out-File -FilePath $logFilePath -Force

# Load blank policy configuration
$PGConfig = Get-DefendpointSettings -LocalFile -FileLocation "$baseFolder\generated_appGroup.xml"

# Define a list to track existing policy names
$existingPolicies = @($PGConfig.Policies | Select-Object -ExpandProperty Name)

# Read and process CSV
$line = 0
Import-Csv -Path $reportFile -Delimiter "," | ForEach-Object {
    $PolicyType = $_."Policy Type"
    $PolicyName = $_."Policy Name"
    $PolicyDescription = ($_."Policy Description" -replace "`r`n|`n|`r", " - ")
    $Active = $_."Active" -eq "No" # Policy is activated or not - if No, then the Policy.Disabled is set to true
    $Action = $_."Action" # Action to be taken to the application
    $SeecurityToken = $_."Security Token" # Token type to be assigned to the application

    # only take Advanced Policy
    if($PolicyType -ne "Advanced Policy") {
        return
    }
    # Validate the policy name
    if (-not $PolicyName -or ($PolicyName -match "macOS" -or $PolicyName -match "Default MAC Policy" -or $PolicyName -match "MAC OS")) {
        Add-Content -Path $logFilePath -Value "Line $line : Policy Name is missing or contains an invalid value ('macOS' or 'Default MAC Policy')."
        return
    }

    # Check for duplicate policy name
    if ($existingPolicies -contains $PolicyName) {
        Add-Content -Path $logFilePath -Value "Line $line : Policy '$PolicyName' already exists in the configuration. Skipping."
        return
    }

    # Log progress
    Write-Host "Processing line $line : $PolicyName"

    # Create a new policy
    $newPolicy = New-Object Avecto.Defendpoint.Settings.Policy($PGConfig)
    $newPolicy.Name = $PolicyName
    $newPolicy.Description = $PolicyDescription
    $newPolicy.Disabled = $Active #take the value of action

    # Find matching Application Group object
    $appGroup = $PGConfig.ApplicationGroups | Where-Object { $_.Name -eq $PolicyName }
    if ($appGroup -ne $null) {
        # Create an application assignment
        $appAssignment = New-Object Avecto.Defendpoint.Settings.ApplicationAssignment($PGConfig)
        if($Action -eq "Block") {
            $appAssignment.Action = "Block"
        } elseif ($Action -eq "Run Normally") {
            $appAssignment.Action = "Allow"
            $appAssignment.TokenType = "Unmodified"
        }else {
            if($SeecurityToken -eq "Administrator") {
                $appAssignment.Action = "Allow"
                $appAssignment.TokenType = "AddAdmin"
            } else {
                $appAssignment.Action = "Allow"
                $appAssignment.TokenType = "Unmodified"
            }
        }
        $appAssignment.Audit = "On"
        $appAssignment.PrivilegeMonitoring = "On"
        $appAssignment.ApplicationGroup = $appGroup

        # Add the application assignment to the policy
        $newPolicy.ApplicationAssignments.Add($appAssignment)
    } else {
        Add-Content -Path $logFilePath -Value "Line $line : No matching application group found for policy '$PolicyName'."
    }

    # Add the new policy to the configuration
    $PGConfig.Policies.Add($newPolicy)

    # Add the policy name to the tracking list
    $existingPolicies += $PolicyName

    # Increment line counter
    $line++
}

# Save the updated configuration to an XML file
Set-DefendpointSettings -SettingsObject $PGConfig -LocalFile -FileLocation $outputFile

Write-Host "Policies with application assignments created and saved to $outputFile"
