param(
    [Parameter(Mandatory=$true)][string]$UserPrincipalName,
    [ValidateSet("ServiceDesk","Engineer")][string]$RolePack = "ServiceDesk",
    [string]$RequestId = ("JOIN-" + (Get-Date -Format "yyyyMMddHHmmss")),
    [switch]$IncludeClientA
)
. "$PSScriptRoot\00_Config.ps1"
Assert-LabTenant

$user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'"
if (-not $user) { throw "User not found: $UserPrincipalName" }

$groups = @($JmlLab.AllStaffGroup)
if ($RolePack -eq "ServiceDesk") { $groups += $JmlLab.ServiceDeskGroup }
if ($RolePack -eq "Engineer") { $groups += $JmlLab.EngineersGroup }
if ($IncludeClientA) { $groups += $JmlLab.ClientAGroup }

$assigned = @()
foreach ($gname in $groups) {
    $gid = Get-GroupIdByName $gname
    try {
        New-MgGroupMember -GroupId $gid -DirectoryObjectId $user.Id -ErrorAction Stop
        $assigned += $gname
    } catch {
        # already member
        $assigned += "$gname (already)"
    }
}

# Ensure account enabled
Update-MgUser -UserId $user.Id -AccountEnabled:$true

$evidence = [pscustomobject]@{
    RequestId = $RequestId
    Action = "Joiner"
    User = $UserPrincipalName
    RolePack = $RolePack
    Groups = $assigned
    TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
    Actor = (Get-MgContext).Account
}
Write-Evidence -Name "Joiner_$($UserPrincipalName.Split('@')[0])" -Data $evidence
$evidence | Format-List
