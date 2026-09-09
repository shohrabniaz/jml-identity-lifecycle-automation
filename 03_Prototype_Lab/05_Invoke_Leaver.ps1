param(
    [Parameter(Mandatory=$true)][string]$UserPrincipalName,
    [string]$RequestId = ("LEAVE-" + (Get-Date -Format "yyyyMMddHHmmss")),
    [switch]$Immediate
)
. "$PSScriptRoot\00_Config.ps1"
Assert-LabTenant

$user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'"
if (-not $user) { throw "User not found: $UserPrincipalName" }

$removeGroups = @(
    $JmlLab.PrivilegedGroup,
    $JmlLab.ServiceDeskGroup,
    $JmlLab.EngineersGroup,
    $JmlLab.ClientAGroup,
    $JmlLab.ClientBGroup,
    $JmlLab.AllStaffGroup
)

$removed = @()
foreach ($gname in $removeGroups) {
    try {
        $gid = Get-GroupIdByName $gname
        Remove-MgGroupMemberByRef -GroupId $gid -DirectoryObjectId $user.Id -ErrorAction Stop
        $removed += $gname
    } catch {
        $removed += "$gname (not member/error ignored)"
    }
}

Update-MgUser -UserId $user.Id -AccountEnabled:$false
# Revoke refresh tokens / sessions where supported
try { Revoke-MgUserSignInSession -UserId $user.Id | Out-Null } catch { }

$evidence = [pscustomobject]@{
    RequestId = $RequestId
    Action = "Leaver"
    Immediate = [bool]$Immediate
    User = $UserPrincipalName
    GroupsRemovedAttempted = $removed
    AccountEnabled = $false
    TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
    Actor = (Get-MgContext).Account
}
Write-Evidence -Name "Leaver_$($UserPrincipalName.Split('@')[0])" -Data $evidence
$evidence | Format-List
