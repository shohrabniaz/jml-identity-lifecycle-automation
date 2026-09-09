# Run full Project 3 lab - Entra ID Free only (no paid Azure resources)
param(
    [switch]$SkipConnect,
    [switch]$UseDeviceAuth
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

. "$PSScriptRoot\Connect-Lab.ps1"

if (-not $SkipConnect) {
    Connect-JmlLabTenant -UseDeviceAuth:$UseDeviceAuth
}

Import-Module Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.SignIns -ErrorAction Stop

$domain = "niazkhanghostgmail.onmicrosoft.com"
$root = $PSScriptRoot

Write-Host ""
Write-Host "=== STEP 1: Create groups ===" -ForegroundColor Cyan
& "$root\01_Create_Lab_Groups.ps1"

Write-Host ""
Write-Host "=== STEP 2: Create synthetic users ===" -ForegroundColor Cyan
& "$root\02_Create_Synthetic_Users.ps1"

Write-Host ""
Write-Host "=== STEP 3: Joiner (ServiceDesk) ===" -ForegroundColor Cyan
& "$root\03_Invoke_Joiner.ps1" -UserPrincipalName "joiner.test01@$domain" -RolePack ServiceDesk

Write-Host ""
Write-Host "=== STEP 4: Joiner (Engineer + ClientA) ===" -ForegroundColor Cyan
& "$root\03_Invoke_Joiner.ps1" -UserPrincipalName "joiner.test02@$domain" -RolePack Engineer -IncludeClientA

Write-Host ""
Write-Host "=== STEP 5: Mover (ServiceDesk to Engineer) ===" -ForegroundColor Cyan
& "$root\04_Invoke_Mover.ps1" -UserPrincipalName "mover.test01@$domain" -NewRolePack Engineer

Write-Host ""
Write-Host "=== STEP 6: Grant privileged (as-is gap simulation) ===" -ForegroundColor Cyan
. "$root\00_Config.ps1"
$privId = Get-GroupIdByName $JmlLab.PrivilegedGroup
$leaver2 = Get-MgUser -Filter "userPrincipalName eq 'leaver.test02@$domain'"
New-MgGroupMember -GroupId $privId -DirectoryObjectId $leaver2.Id -ErrorAction SilentlyContinue | Out-Null
Write-Host "Added leaver.test02 to $($JmlLab.PrivilegedGroup)"

Write-Host ""
Write-Host "=== STEP 7: Leaver (automated - leaver.test01) ===" -ForegroundColor Cyan
& "$root\05_Invoke_Leaver.ps1" -UserPrincipalName "leaver.test01@$domain" -Immediate

Write-Host ""
Write-Host "=== STEP 8: Simulate failed manual leaver (orphan risk) ===" -ForegroundColor Cyan
Update-MgUser -UserId $leaver2.Id -AccountEnabled:$false
try { Revoke-MgUserSignInSession -UserId $leaver2.Id | Out-Null } catch {}
Write-Evidence -Name "FailedLeaver_OrphanSim" -Data @{
    User = "leaver.test02@$domain"
    Note = "Account disabled but privileged group membership retained (as-is failure mode)"
    TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
}

Write-Host ""
Write-Host "=== STEP 9: Orphan privileged report ===" -ForegroundColor Cyan
& "$root\06_Orphan_Privileged_Report.ps1"

Write-Host ""
Write-Host "=== LAB COMPLETE ===" -ForegroundColor Green
Write-Host "Evidence JSON files: $root\evidence"
Write-Host "Cost: Entra ID Free only - no VMs, no M365 licenses, no Pay-as-you-go."
