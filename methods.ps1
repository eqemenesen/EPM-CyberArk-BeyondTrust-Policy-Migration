Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\Avecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Cmdlets.dll' -Force -ErrorAction Stop
Import-Module 'C:\Program Files\Avecto\Privilege Guard Management Consoles\PowerShell\A.vecto.Defendpoint.Cmdlets\Avecto.Defendpoint.Settings.dll' -Force -ErrorAction Stop

# Initialize AccountsFilter
$accountsFilter = New-Object Avecto.Defendpoint.Settings.AccountsFilter
$accountsFilter.InverseFilter = $false
$accountsFilter.Accounts = New-Object Avecto.Defendpoint.Settings.AccountList
$accountsFilter.Accounts.WindowsAccounts = New-Object System.Collections.Generic.List[Avecto.Defendpoint.Settings.Account]

# Create Account Object
$UserAccount = New-Object Avecto.Defendpoint.Settings.Account
$UserAccount.Name = "DIJITALVARLIK\USR_SRV_DAS_GTSRSServiceControlCenter"
$UserAccount.Group = $false

# Add to AccountsFilter
$accountsFilter.Accounts.WindowsAccounts.Add($UserAccount)





                                            
$appAssignment = New-Object Avecto.Defendpoint.Settings.ApplicationAssignment($PGConfig)

# See what members it actually has
$appAssignment | Get-Member -Force
