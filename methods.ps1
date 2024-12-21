# Change this to the path where you want to save the output
$outputFilePath = 'C:\Temp\Avecto_Defendpoint_TypesAndMembers.txt'

# Wrap your script in a script block and pipe it to Out-File
& {
    # 1) Load the assemblies
    $cmdletsAssemblyPath  = 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll'
    $settingsAssemblyPath = 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll'

    [System.Reflection.Assembly]::LoadFile($cmdletsAssemblyPath)  | Out-Null
    [System.Reflection.Assembly]::LoadFile($settingsAssemblyPath) | Out-Null

    # 2) Retrieve .NET types
    $cmdletsAssembly  = [System.Reflection.Assembly]::LoadFile($cmdletsAssemblyPath)
    $settingsAssembly = [System.Reflection.Assembly]::LoadFile($settingsAssemblyPath)

    $cmdletsTypes  = $cmdletsAssembly.GetTypes()
    $settingsTypes = $settingsAssembly.GetTypes()

    # 3) Dump all members for each type in the Cmdlets assembly
    foreach ($type in $cmdletsTypes) {
        Write-Output "===== Type: $($type.FullName) ====="
        $type | Get-Member -Force | Format-Table Name, MemberType, Definition -AutoSize
        Write-Output "`n"
    }

    # 4) Dump all members for each type in the Settings assembly
    foreach ($type in $settingsTypes) {
        Write-Output "===== Type: $($type.FullName) ====="
        $type | Get-Member -Force | Format-Table Name, MemberType, Definition -AutoSize
        Write-Output "`n"
    }
} | Out-File -FilePath $outputFilePath -Encoding UTF8

Write-Host "Reflection output saved to $outputFilePath"
