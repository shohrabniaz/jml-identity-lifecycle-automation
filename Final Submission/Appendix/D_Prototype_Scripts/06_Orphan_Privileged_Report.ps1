param()
. "$PSScriptRoot\00_Config.ps1"
Assert-LabTenant

$privId = Get-GroupIdByName $JmlLab.PrivilegedGroup
$members = Get-MgGroupMember -GroupId $privId -All
$report = @()
foreach ($m in $members) {
    $u = Get-MgUser -UserId $m.Id -Property Id,DisplayName,UserPrincipalName,AccountEnabled,Department,JobTitle -ErrorAction SilentlyContinue
    if ($u) {
        $report += [pscustomobject]@{
            DisplayName = $u.DisplayName
            UPN = $u.UserPrincipalName
            AccountEnabled = $u.AccountEnabled
            Department = $u.Department
            JobTitle = $u.JobTitle
            Risk = if (-not $u.AccountEnabled) { "DISABLED_BUT_STILL_IN_PRIV_GROUP" } else { "ACTIVE_PRIVILEGED" }
        }
    }
}
Write-Evidence -Name "OrphanPrivilegedReport" -Data $report
$report | Format-Table -AutoSize
