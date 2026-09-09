# PRIVATE CONFIG - do not submit secrets
# Copy values from 07_Lab_Tenant_Setup/Lab_Credentials_PRIVATE.txt
$Global:JmlLab = @{
    TenantId          = "a76a68a8-6243-4b0e-b11c-98e815e91820"
    Domain            = "niazkhanghostgmail.onmicrosoft.com"
    EvidencePath      = ".\evidence"
    PrivilegedGroup   = "SG-CMIT-Privileged-Admins"
    AllStaffGroup     = "SG-CMIT-AllStaff"
    ServiceDeskGroup  = "SG-CMIT-ServiceDesk"
    EngineersGroup    = "SG-CMIT-Engineers"
    ClientAGroup      = "SG-ClientA-Support"
    ClientBGroup      = "SG-ClientB-Support"
}

New-Item -ItemType Directory -Force -Path $Global:JmlLab.EvidencePath | Out-Null

function Assert-LabTenant {
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Not connected. Run Connect-MgGraph first." }
    if ($Global:JmlLab.TenantId -eq "PASTE-TENANT-ID") { throw "Set TenantId in 00_Config.ps1" }
    if ($ctx.TenantId -ne $Global:JmlLab.TenantId) {
        throw "Connected tenant $($ctx.TenantId) != configured lab $($Global:JmlLab.TenantId). Aborting."
    }
}

function Write-Evidence {
    param([string]$Name,[object]$Data)
    $file = Join-Path $Global:JmlLab.EvidencePath ("{0}_{1}.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"), $Name)
    $Data | ConvertTo-Json -Depth 6 | Set-Content -Path $file -Encoding UTF8
    Write-Host "Evidence written: $file"
}

function Get-GroupIdByName {
    param([string]$DisplayName)
    $g = Get-MgGroup -Filter "displayName eq '$DisplayName'"
    if (-not $g) { throw "Group not found: $DisplayName" }
    return $g.Id
}
