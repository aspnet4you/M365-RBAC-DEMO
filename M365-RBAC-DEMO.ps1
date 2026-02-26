<#
    Granular M365 RBAC for Exchange & SharePoint Using Entra ID, PowerShell, 
    and Attribute‑Based Scopes with Mail-Enabled Security Group. 
#>

# -----------------------------
# CONFIGURATIONS. Change them based on your environment
# -----------------------------
$TenantId          = "c33386cf-6e11-484c-a983-b49975ce571a"
$AppDisplayName    = "M365-RBAC-DEMO"
$MailboxAttribute  = "CustomAttribute1"
$MailboxAttrValue  = $AppDisplayName
$MgmtScopeName     = "$AppDisplayName-Attribute-Scope"
$RoleAssignmentName= "$AppDisplayName-Role-Assignment"
$MESG              = "$AppDisplayName-MESG"
$MESGAlias         = "M365RBACDemoMESG"
$member1           = "Paul.Smith@aspnet4you2.onmicrosoft.com"
$member2           = "Bob.Smith@aspnet4you2.onmicrosoft.com"
$spsite            = "https://graph.microsoft.com/v1.0/sites/aspnet4you2.sharepoint.com:/sites/Graph-Demo"

# -----------------------------
# CONNECT TO GRAPH
# -----------------------------
Connect-MgGraph -TenantId $TenantId -Scopes @(
    "Application.ReadWrite.All",
    "Directory.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "Sites.FullControl.All"
)

# -----------------------------
# CREATE APP + SERVICE PRINCIPAL
# -----------------------------
$app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience "AzureADMyOrg"
Write-Host "New App Created - AppDisplayName is $AppDisplayName and AppId is $app.AppId" -ForegroundColor Green

$sp  = New-MgServicePrincipal -AppId $app.AppId
Write-Host "New Service Principal Created - Service Principal Id is $sp.Id" -ForegroundColor Green


# -----------------------------
# You can use this section, if you want to use 
# an existing Entra ID application
# -----------------------------

# $app = Get-MgApplication | Where-Object {$_.DisplayName -eq $AppDisplayName}
# $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"

# -----------------------------
# Get Graph Service Principal (Microsoft Graph)
# -----------------------------
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# -----------------------------
# ASSIGN GRAPH Mail.Read, Sites.Selected, Sites.Read.All (APP PERMISSION)
# I commented Mail.Read because doing so will allow the app to all mailboxes
# I commented Sites.Read.All because doing so will allow the app to all SharePoint sites
# -----------------------------

#
# $mailRead = $graphSp.AppRoles | Where-Object {
#    $_.Value -eq "Mail.Read" -and $_.AllowedMemberTypes -contains "Application"
# }

$sitesSelected = $graphSp.AppRoles | Where-Object {
    $_.Value -eq "Sites.Selected" -and $_.AllowedMemberTypes -contains "Application"
}

#
# $sitesReadAll = $graphSp.AppRoles | Where-Object {
#    $_.Value -eq "Sites.Read.All" -and $_.AllowedMemberTypes -contains "Application"
# }

# New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id `
#    -PrincipalId $sp.Id `
#    -ResourceId $graphSp.Id `
#    -AppRoleId $mailRead.Id

New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id `
    -PrincipalId $sp.Id `
    -ResourceId $graphSp.Id `
    -AppRoleId $sitesSelected.Id

# New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id `
#    -PrincipalId $sp.Id `
#    -ResourceId $graphSp.Id `
#    -AppRoleId $sitesReadAll.Id

# -----------------------------
# ASSIGN Delegated Mail.Read, Sites.ReadWrite.All to the app registration
# -----------------------------

$mailReadDelegated = $graphSp.Oauth2PermissionScopes | Where-Object {
    $_.Value -eq "Mail.Read" -and $_.Type -eq "User"
}

$sitesReadWriteAllDelegated = $graphSp.Oauth2PermissionScopes | Where-Object { 
    $_.Value -eq "Sites.ReadWrite.All" -and $_.Type -eq "User"
}

Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
        ResourceAccess = @(
            @{
                Id   = $mailReadDelegated.Id
                Type = "Scope"   # Delegated permission
            }
            @{
                Id   = $sitesReadWriteAllDelegated.Id
                Type = "Scope"   # Delegated permission
            }
        )
    }
)

# -----------------------------
# Grant admin consent for scopes to the app registration 
# -----------------------------

New-MgOauth2PermissionGrant -BodyParameter @{
    clientId     = $sp.Id          # client = your app's service principal
    consentType  = "AllPrincipals"
    principalId  = $null
    resourceId   = $graphSp.Id     # resource = Microsoft Graph SP
    scope        = "Mail.Read Sites.ReadWrite.All"
}


# -----------------------------
# CONNECT TO EXCHANGE ONLINE
# YOU MAY WANT TO RUN THIS COMMAND AS YOUR FIRST LINE OF
# EXECUTION TO AVOID WAM ERROR.
# -----------------------------
Connect-ExchangeOnline -DisableWAM


# -----------------------------
# Create Mail-enabled security group and assign members
# -----------------------------

New-DistributionGroup `
    -Name $MESG `
    -DisplayName $MESG `
    -Alias $MESGAlias `
    -Type Security

$newmembers = @(
    $member1,
    $member2
)

foreach ($m in $newmembers) {
    Add-DistributionGroupMember -Identity $MESG -Member $m
}


# -----------------------------
# SYNC MESG MEMBERSHIP → ATTRIBUTE
# -----------------------------
$members = Get-DistributionGroupMember $MESG

foreach ($m in $members) {
    if ($m.RecipientType -eq "UserMailbox") {
        
        $setParams = @{
            Identity = $m.PrimarySmtpAddress
        }

        $setParams[$MailboxAttribute] = $MailboxAttrValue

        Set-Mailbox @setParams
    }
}

# -----------------------------
# CREATE ATTRIBUTE-BASED SCOPE
# -----------------------------
$scopeFilter = "$MailboxAttribute -eq '$MailboxAttrValue'"

$existingScope = Get-ManagementScope -ErrorAction SilentlyContinue |
    Where-Object {$_.Name -eq $MgmtScopeName}

if (-not $existingScope) {
    $mgmtScope = New-ManagementScope -Name $MgmtScopeName -RecipientRestrictionFilter $scopeFilter
} else {
    $mgmtScope = $existingScope
}

# -----------------------------
# CREATE EXO SERVICE PRINCIPAL POINTER
# -----------------------------
#$exoSp = New-ServicePrincipal -AppId $app.AppId -DisplayName $AppDisplayName -ErrorAction SilentlyContinue
$exoSp = New-ServicePrincipal -AppId $app.AppId -ObjectId $sp.Id -DisplayName $AppDisplayName -ErrorAction SilentlyContinue

if (-not $exoSp) {
    $exoSp = Get-ServicePrincipal | Where-Object { $_.AppId -eq $app.AppId }
}

# -----------------------------
# ASSIGN APPLICATION Mail.Read ROLE WITH ATTRIBUTE SCOPE
# -----------------------------
$roleName = "Application Mail.Read"

$existingAssignment = Get-ManagementRoleAssignment -ErrorAction SilentlyContinue |
    Where-Object {$_.Name -eq $RoleAssignmentName}

if (-not $existingAssignment) {
    New-ManagementRoleAssignment `
        -Name $RoleAssignmentName `
        -Role $roleName `
        -App $exoSp.Id `
        -CustomResourceScope $mgmtScope.Identity
}

Write-Host "Setup complete. Service principal is scoped to MESG members via attribute." -ForegroundColor Green

Get-ManagementRoleAssignment -Identity $RoleAssignmentName
Get-ManagementScope -Identity $MgmtScopeName


# -----------------------------
# Grant the service principal permission at the SharePoint Site.
# Remember this permission is for your app (not for delegated user).
# Delegated permissions are defined above.
# -----------------------------

$site = Invoke-MgGraphRequest -Method GET `
    -Uri $spsite
$siteId = $site.id

$body = @{
    roles = @("fullcontrol")
    grantedToIdentities = @(
        @{
            application = @{
                id = $app.AppId
                displayName = $AppDisplayName
            }
        }
    )
}


Invoke-MgGraphRequest -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" `
    -Body ($body | ConvertTo-Json -Depth 5)


$permissions = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"