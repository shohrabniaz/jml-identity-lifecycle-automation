param()
. "$PSScriptRoot\00_Config.ps1"
Assert-LabTenant

$users = @(
    @{ Given="Joiner"; Surname="Test01"; UPN="joiner.test01@$($JmlLab.Domain)"; Dept="ServiceDesk"; Title="Service Desk Analyst" },
    @{ Given="Joiner"; Surname="Test02"; UPN="joiner.test02@$($JmlLab.Domain)"; Dept="Engineers"; Title="Cloud Engineer" },
    @{ Given="Mover";  Surname="Test01"; UPN="mover.test01@$($JmlLab.Domain)"; Dept="ServiceDesk"; Title="Service Desk Analyst" },
    @{ Given="Leaver"; Surname="Test01"; UPN="leaver.test01@$($JmlLab.Domain)"; Dept="Engineers"; Title="Senior Engineer" },
    @{ Given="Leaver"; Surname="Test02"; UPN="leaver.test02@$($JmlLab.Domain)"; Dept="Engineers"; Title="Privileged Admin" }
)

$passwordProfile = @{
    ForceChangePasswordNextSignIn = $true
    Password = "ChangeMe!LabOnly#$((Get-Random -Maximum 99999))"
}

$results = @()
foreach ($u in $users) {
    $existing = Get-MgUser -Filter "userPrincipalName eq '$($u.UPN)'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "User exists: $($u.UPN)"
        $results += $existing
        continue
    }
    $body = @{
        AccountEnabled = $true
        DisplayName = "$($u.Given) $($u.Surname)"
        MailNickname = ($u.UPN.Split("@")[0] -replace "[^a-zA-Z0-9]", "")
        UserPrincipalName = $u.UPN
        GivenName = $u.Given
        Surname = $u.Surname
        Department = $u.Dept
        JobTitle = $u.Title
        PasswordProfile = $passwordProfile
    }
    $new = New-MgUser -BodyParameter $body
    Write-Host "Created $($u.UPN)"
    $results += $new
}
Write-Evidence -Name "CreateUsers" -Data ($results | Select-Object Id, DisplayName, UserPrincipalName, Department, JobTitle)
Write-Host "NOTE: Initial passwords are random per run. Force change on next sign-in is enabled. Do not store passwords in dissertation."
