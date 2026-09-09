# Export lab state as text evidence files for evaluation screenshots (C8)
# Run while signed in: . .\Connect-Lab.ps1; Connect-JmlLabTenant
param()

$ErrorActionPreference = "Stop"
$out = Join-Path $PSScriptRoot "..\05_Evaluation\screenshots"
New-Item -ItemType Directory -Force -Path $out | Out-Null

. "$PSScriptRoot\Connect-Lab.ps1"
. "$PSScriptRoot\00_Config.ps1"

$ctx = Get-MgContext
if (-not $ctx) { Connect-JmlLabTenant }
Assert-LabTenant
Import-Module Microsoft.Graph.Users, Microsoft.Graph.Groups -ErrorAction Stop

$domain = $JmlLab.Domain

Write-Host "Exporting lab evidence to $out" -ForegroundColor Cyan

# 01 Users
$users = Get-MgUser -Filter "startswith(userPrincipalName,'joiner.') or startswith(userPrincipalName,'mover.') or startswith(userPrincipalName,'leaver.')" -Property DisplayName,UserPrincipalName,AccountEnabled,JobTitle
$users | Select-Object DisplayName, UserPrincipalName, AccountEnabled, JobTitle | Format-Table -AutoSize | Out-String | Set-Content (Join-Path $out "01_Entra_Users_List.txt") -Encoding UTF8

# 02 Groups
$groups = Get-MgGroup -Filter "startswith(displayName,'SG-CMIT') or startswith(displayName,'SG-Client') or startswith(displayName,'SG-Role')" -Property DisplayName,SecurityEnabled
$groups | Select-Object DisplayName, SecurityEnabled | Format-Table -AutoSize | Out-String | Set-Content (Join-Path $out "02_Entra_Groups_List.txt") -Encoding UTF8

# 03 Joiner memberships
$joiner = Get-MgUser -Filter "userPrincipalName eq 'joiner.test01@$domain'"
$memberOf = Get-MgUserMemberOf -UserId $joiner.Id -All
$names = foreach ($m in $memberOf) {
    if ($m.AdditionalProperties.displayName) { $m.AdditionalProperties.displayName }
}
@(
    "User: $($joiner.UserPrincipalName)"
    "DisplayName: $($joiner.DisplayName)"
    "Group memberships:"
    ($names | ForEach-Object { "  - $_" })
) | Set-Content (Join-Path $out "03_Joiner_Group_Memberships.txt") -Encoding UTF8

# 04 Leaver disabled
$leaver1 = Get-MgUser -Filter "userPrincipalName eq 'leaver.test01@$domain'" -Property DisplayName,UserPrincipalName,AccountEnabled,Department
@(
    "User: $($leaver1.UserPrincipalName)"
    "DisplayName: $($leaver1.DisplayName)"
    "AccountEnabled: $($leaver1.AccountEnabled)"
    "Expected: False (automated leaver)"
) | Set-Content (Join-Path $out "04_Leaver_Disabled_Account.txt") -Encoding UTF8

# 05 Privileged group members
$privId = Get-GroupIdByName $JmlLab.PrivilegedGroup
$members = Get-MgGroupMember -GroupId $privId -All
$rows = @()
foreach ($m in $members) {
    $u = Get-MgUser -UserId $m.Id -Property DisplayName,UserPrincipalName,AccountEnabled -ErrorAction SilentlyContinue
    if ($u) { $rows += $u }
}
@(
    "Group: $($JmlLab.PrivilegedGroup)"
    "Members:"
    ($rows | ForEach-Object { "  $($_.DisplayName) | $($_.UserPrincipalName) | Enabled=$($_.AccountEnabled)" })
) | Set-Content (Join-Path $out "05_Privileged_Group_Members.txt") -Encoding UTF8

# 06 Orphan report
& "$PSScriptRoot\06_Orphan_Privileged_Report.ps1" *>&1 | Out-String | Set-Content (Join-Path $out "06_PowerShell_Orphan_Report.txt") -Encoding UTF8

Write-Host "Done. Open .txt files and snip with Win+Shift+S, or run: python ..\_render_screenshot_pngs.py" -ForegroundColor Green
