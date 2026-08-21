<#
.SYNOPSIS
  One-time, admin-only bootstrap for a new practice site. Creates a Microsoft 365
  group-connected team site (an EXACT mirror of the Gift & Estate site, which is a
  GROUP#0 team site) and grants the ProjectSync certificate app a Sites.Selected
  fullControl grant on it -- the two things the cert-authed ProvisionPracticeSite
  tool cannot do itself.

.DESCRIPTION
  You run this interactively with YOUR admin identity. It uses only the Microsoft Graph
  PowerShell module (Microsoft.Graph.Authentication), whose sign-in app is a Microsoft
  first-party app already consented in the tenant -- so NO app registration is required.
  PS 5.1-compatible (pure ASCII, no PnP dependency).

  Creating the M365 group auto-provisions its SharePoint team site at
  https://<tenant>.sharepoint.com/sites/<mailNickname>. The <mailNickname> is taken from
  the last segment of -SiteUrl, so pass the URL you want and it drives the nickname.

  NOTE: creating an M365 group requires the signed-in user to be allowed to create groups
  (Groups/User Administrator or Global Admin, or an unrestricted group-creation policy) and
  the Graph scopes below to be consented. Graph also auto-adds the SIGNED-IN user as an
  owner+member of the new group in addition to -OwnerUpn; run it as the intended owner
  (or prune the extra owner afterward) to keep a single owner.

  After it finishes, run:
    dotnet run --project tools/ProvisionPracticeSite -- `
      --practice "Marital Dissolution" --leader rhoffman@marshall-stevens.com `
      --to <SiteUrl> --apply

.PARAMETER SiteUrl
  Full URL the new team site should live at, e.g.
  https://marshallstevens.sharepoint.com/sites/MaritalDissolution
  (its last path segment becomes the group mailNickname).

.PARAMETER Title
  Group + site display name, e.g. "Marital Dissolution".

.PARAMETER OwnerUpn
  The person who should own the group/site (the sole Owner, so nobody can accidentally
  delete it). Prompted for if omitted.

.PARAMETER Visibility
  M365 group visibility: Private (default) or Public.

.PARAMETER AppClientId
  The ProjectSync cert app's client id to grant. Default is the current app
  (0c553a39-81c7-48aa-9a06-a9e9c132884b) -- SharePoint:ClientId in local.settings.json.

.PARAMETER AppDisplayName
  Display name recorded on the Sites.Selected grant (cosmetic).

.PARAMETER WhatIf
  Show what would happen without creating the group or granting anything.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter(Mandatory = $true)] [string] $Title,
    [string] $OwnerUpn,
    [ValidateSet("Private", "Public")] [string] $Visibility = "Private",
    [string] $AppClientId = "0c553a39-81c7-48aa-9a06-a9e9c132884b",
    [string] $AppDisplayName = "ProjectSync",
    [switch] $DeviceCode
)

$ErrorActionPreference = "Stop"

$uri          = [Uri]$SiteUrl
$hostName     = $uri.Host                                   # marshallstevens.sharepoint.com
$mailNickname = ($uri.AbsolutePath.Trim('/') -split '/')[-1] # MaritalDissolution

Write-Host "=== Bootstrap practice site (M365 group team site, mirrors Gift & Estate) ===" -ForegroundColor Cyan
Write-Host "  Site URL     : $SiteUrl"
Write-Host "  Title        : $Title"
Write-Host "  mailNickname : $mailNickname"
Write-Host "  Visibility   : $Visibility"
Write-Host "  Grant app    : $AppClientId ($AppDisplayName)"
Write-Host ""

# ---------------------------------------------------------------------------
# Ensure the Graph module is present (current user scope; no admin install).
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication")) {
    Write-Host "Installing module Microsoft.Graph.Authentication (CurrentUser)..." -ForegroundColor Yellow
    Install-Module "Microsoft.Graph.Authentication" -Scope CurrentUser -Force -AllowClobber
}
Import-Module "Microsoft.Graph.Authentication" -ErrorAction Stop

if (-not $OwnerUpn) {
    $OwnerUpn = Read-Host "Primary owner UPN for the new site (e.g. jparks@marshall-stevens.com)"
    if (-not $OwnerUpn) { throw "Owner UPN is required. Re-run with -OwnerUpn <upn>." }
}

# Group creation, owner resolution, and the Sites.Selected grant are all Graph calls.
# WAM's interactive window opens HIDDEN behind embedded/IDE terminals (shows as "user canceled"), so use
# -DeviceCode there: it prints a URL + code in the console instead of popping a window.
$connect = @{
    Scopes    = @("Group.ReadWrite.All", "User.Read.All", "Sites.FullControl.All")
    NoWelcome = $true
}
if ($DeviceCode) { $connect["UseDeviceAuthentication"] = $true }
Connect-MgGraph @connect

function Graph-Get([string]$Path) {
    return Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0$Path" -OutputType PSObject
}

# ---------------------------------------------------------------------------
# STEP 1 - create the M365 group (auto-provisions the mirrored team site).
# ---------------------------------------------------------------------------
Write-Host "STEP 1: create M365 group + team site" -ForegroundColor Cyan

$owner = Graph-Get "/users/$OwnerUpn`?`$select=id,displayName,userPrincipalName"
$ownerId = $owner.id
Write-Host "  Owner resolved: $($owner.displayName) <$($owner.userPrincipalName)> ($ownerId)"

# Idempotency: reuse an existing group with this mailNickname if present.
$found = (Graph-Get "/groups`?`$filter=mailNickname eq '$mailNickname'&`$select=id,displayName,mailNickname").value
$group = $null
if ($found -and $found.Count -gt 0) {
    $group = $found[0]
    Write-Host "  [OK] group '$($group.displayName)' already exists ($($group.id)) - reusing." -ForegroundColor Green
}
elseif ($PSCmdlet.ShouldProcess("$Title ($mailNickname)", "create M365 group")) {
    $body = @{
        displayName     = $Title
        mailNickname    = $mailNickname
        mailEnabled     = $true
        securityEnabled = $false
        groupTypes      = @("Unified")
        visibility      = $Visibility
        "owners@odata.bind" = @("https://graph.microsoft.com/v1.0/users/$ownerId")
    } | ConvertTo-Json -Depth 6

    $group = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" `
        -Body $body -ContentType "application/json" -OutputType PSObject
    Write-Host "  [OK] group created ($($group.id))." -ForegroundColor Green
}
else {
    Write-Host "  (WhatIf) would create the group; cannot continue without it." -ForegroundColor Yellow
    return
}

$groupId = $group.id

# The SharePoint team site provisions asynchronously behind the group - poll for it.
Write-Host "  waiting for the SharePoint team site to provision..."
$site = $null
for ($i = 0; $i -lt 30 -and -not $site; $i++) {
    Start-Sleep -Seconds 10
    try { $site = Graph-Get "/groups/$groupId/sites/root`?`$select=id,webUrl" } catch { $site = $null }
    if ($site) { Write-Host "    provisioned: $($site.webUrl)" }
    else { Write-Host "    still provisioning... ($(($i + 1) * 10)s)" }
}
if (-not $site) { throw "Team site did not provision within ~5 minutes. Check the group in the M365 admin center, then re-run (idempotent)." }

$graphSiteId = $site.id
if ($site.webUrl.TrimEnd('/') -ne $SiteUrl.TrimEnd('/')) {
    Write-Host "  WARNING: site provisioned at $($site.webUrl), not the requested $SiteUrl" -ForegroundColor Yellow
    Write-Host "           (mailNickname collision?). Use the ACTUAL url above for --to and the config." -ForegroundColor Yellow
}
Write-Host ""

# ---------------------------------------------------------------------------
# STEP 2 - grant the cert app Sites.Selected fullControl on the new site (Graph).
# Same approach used to grant the app on Gift & Estate.
# ---------------------------------------------------------------------------
Write-Host "STEP 2: grant $AppDisplayName Sites.Selected fullControl" -ForegroundColor Cyan
Write-Host "  Graph site id: $graphSiteId"

$existingPerms = (Graph-Get "/sites/$graphSiteId/permissions").value
$already = $existingPerms | Where-Object { $_.grantedToIdentitiesV2.application.id -eq $AppClientId }
if ($already) {
    Write-Host "  [OK] app already has a permission grant here (roles: $($already.roles -join ', '))." -ForegroundColor Green
}
elseif ($PSCmdlet.ShouldProcess($site.webUrl, "grant $AppClientId fullControl")) {
    $body = @{
        roles = @("fullControl")
        grantedToIdentities = @(@{
            application = @{ id = $AppClientId; displayName = $AppDisplayName }
        })
    } | ConvertTo-Json -Depth 6
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$graphSiteId/permissions" `
        -Body $body -ContentType "application/json" | Out-Null
    Write-Host "  [OK] granted fullControl." -ForegroundColor Green
}
Write-Host ""

Write-Host "Bootstrap complete. Next:" -ForegroundColor Cyan
Write-Host "  dotnet run --project tools/ProvisionPracticeSite -- ``"
Write-Host "    --practice `"$Title`" --to $($site.webUrl) --apply"
