# Connect to the Project 3 lab tenant (run once per PowerShell session)
$script:JmlLabTenantId = "a76a68a8-6243-4b0e-b11c-98e815e91820"
$script:JmlLabScopes = @(
    "User.ReadWrite.All",
    "Group.ReadWrite.All",
    "Directory.ReadWrite.All",
    "AuditLog.Read.All"
)

function Connect-JmlLabTenant {
    param([switch]$UseDeviceAuth)

    $ProgressPreference = "SilentlyContinue"

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "Installing Microsoft.Graph module (one-time)..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $existing = Get-MgContext -ErrorAction SilentlyContinue
    if ($existing -and $existing.TenantId -eq $script:JmlLabTenantId) {
        Write-Host "Already signed in as $($existing.Account)" -ForegroundColor Green
        return $existing
    }

    if ($existing) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Start-Sleep -Seconds 2
    }

    if ($UseDeviceAuth) {
        Write-Host ""
        Write-Host "DEVICE CODE SIGN-IN" -ForegroundColor Yellow
        Write-Host "1. Open https://login.microsoft.com/device in your browser"
        Write-Host "2. Enter the code shown below (wait for it to appear)"
        Write-Host "3. Sign in with your Azure GLOBAL ADMIN account"
        Write-Host ""
        Connect-MgGraph -TenantId $script:JmlLabTenantId -Scopes $script:JmlLabScopes -UseDeviceAuthentication -NoWelcome -ContextScope CurrentUser
    } else {
        Write-Host ""
        Write-Host "Opening browser for sign-in..." -ForegroundColor Yellow
        Write-Host "Use your Azure GLOBAL ADMIN account for niazkhanghostgmail.onmicrosoft.com"
        Write-Host ""
        Connect-MgGraph -TenantId $script:JmlLabTenantId -Scopes $script:JmlLabScopes -NoWelcome -ContextScope CurrentUser
    }

    $ctx = Get-MgContext
    if (-not $ctx) {
        throw "Sign-in did not complete. Close this window and try again."
    }
    if ($ctx.TenantId -ne $script:JmlLabTenantId) {
        throw "Wrong tenant. Expected $($script:JmlLabTenantId) but got $($ctx.TenantId)."
    }

    Write-Host ""
    Write-Host "Signed in OK: $($ctx.Account)" -ForegroundColor Green
    return $ctx
}
