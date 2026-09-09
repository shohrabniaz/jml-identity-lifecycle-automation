# Run AFTER Azure signup - discovers tenant info for 00_Config.ps1
# Usage: Connect-MgGraph -Scopes "Directory.Read.All","User.Read.All" then run this script

$org = Get-MgOrganization | Select-Object -First 1
$domain = ($org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true }).Name
if (-not $domain) {
    $domain = ($org.VerifiedDomains | Select-Object -First 1).Name
}

Write-Host ""
Write-Host "=== LAB TENANT INFO (copy into 00_Config.ps1) ===" -ForegroundColor Cyan
Write-Host "TenantId : $($org.Id)"
Write-Host "Domain   : $domain"
Write-Host "Name     : $($org.DisplayName)"
Write-Host ""
Write-Host "Suggested org display name: Niaz-SCI700-JML-Lab" -ForegroundColor Yellow
Write-Host "Next: update 00_Config.ps1, then run 01_Create_Lab_Groups.ps1" -ForegroundColor Green
