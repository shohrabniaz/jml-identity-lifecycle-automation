param()
. "$PSScriptRoot\00_Config.ps1"
Assert-LabTenant

$names = @(
    $JmlLab.AllStaffGroup,
    $JmlLab.ServiceDeskGroup,
    $JmlLab.EngineersGroup,
    $JmlLab.PrivilegedGroup,
    $JmlLab.ClientAGroup,
    $JmlLab.ClientBGroup,
    "SG-Role-Joiners-Pending",
    "SG-Role-Leavers-Pending"
)

$created = @()
foreach ($n in $names) {
    $existing = Get-MgGroup -Filter "displayName eq '$n'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Exists: $n"
        $created += $existing
        continue
    }
    $mailNick = ($n -replace "[^a-zA-Z0-9]", "").Substring(0, [Math]::Min(20, ($n -replace "[^a-zA-Z0-9]", "").Length))
    $g = New-MgGroup -DisplayName $n -MailEnabled:$false -MailNickname $mailNick -SecurityEnabled:$true
    Write-Host "Created: $n"
    $created += $g
}
Write-Evidence -Name "CreateGroups" -Data $created
