param(
    [Parameter(Mandatory=$true)][string]$UserPrincipalName,
    [ValidateSet("ServiceDesk","Engineer")][string]$NewRolePack,
    [string]$RequestId = ("MOVE-" + (Get-Date -Format "yyyyMMddHHmmss"))
)
. "$PSScriptRoot\00_Config.ps1"
Assert-LabTenant

$user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'"
if (-not $user) { throw "User not found: $UserPrincipalName" }

$sd = Get-GroupIdByName $JmlLab.ServiceDeskGroup
$eng = Get-GroupIdByName $JmlLab.EngineersGroup

# Remove both role groups then add target pack (simple lab model)
foreach ($gid in @($sd, $eng)) {
    try { Remove-MgGroupMemberByRef -GroupId $gid -DirectoryObjectId $user.Id -ErrorAction SilentlyContinue } catch {}
}

$addName = if ($NewRolePack -eq "ServiceDesk") { $JmlLab.ServiceDeskGroup } else { $JmlLab.EngineersGroup }
$addId = Get-GroupIdByName $addName
New-MgGroupMember -GroupId $addId -DirectoryObjectId $user.Id -ErrorAction SilentlyContinue | Out-Null
Update-MgUser -UserId $user.Id -Department $NewRolePack | Out-Null

$evidence = [pscustomobject]@{
    RequestId = $RequestId
    Action = "Mover"
    User = $UserPrincipalName
    NewRolePack = $NewRolePack
    TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
    Actor = (Get-MgContext).Account
}
Write-Evidence -Name "Mover_$($UserPrincipalName.Split('@')[0])" -Data $evidence
$evidence | Format-List
