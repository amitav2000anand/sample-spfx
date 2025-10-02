# Step 1: Connect to SharePoint using certificate-based authentication
Connect-PnPOnline -Url "https://s63fb.sharepoint.com/sites/selfserviceagent" `
  -ClientId "4ffd1a7a-9a30-48a9-bc3d-51060e46591b" `
  -Tenant "s63fb.onmicrosoft.com" `
  -CertificatePath (Join-Path (Get-Location) "TestCert.pfx") `
  -CertificatePassword (ConvertTo-SecureString "YourStrongPassword123!" -AsPlainText -Force)

Write-Host "`n✅ Connected to SharePoint"


# Step 2: Fetch and display Site-scoped custom actions
# $siteActions = Get-PnPCustomAction -Scope Site
# Write-Host "`n📦 Site-scoped Custom Actions:"
# foreach ($action in $siteActions) {
#     Write-Host "`n🔹 Name: $($action.Name)"
#     Write-Host "🔹 ID: $($action.Id)"
#     Write-Host "🔹 Location: $($action.Location)"
#     Write-Host "🔹 Component ID: $($action.ClientSideComponentId)"
#     Write-Host "🔹 Properties:"
#     if ($action.ClientSideComponentProperties) {
#         $props = $action.ClientSideComponentProperties | ConvertFrom-Json
#         $props | Format-List
#     } else {
#         Write-Host "   (No properties defined)"
#     }
# }

# Step 3: Fetch and display Web-scoped custom actions
$webActions = Get-PnPCustomAction -Scope Web
Write-Host "`n📦 Web-scoped Custom Actions:"
foreach ($action in $webActions) {
    $webActions | Format-List Name, Id, Location, ClientSideComponentId, ClientSideComponentProperties

    # if ($action.ClientSideComponentProperties) {
    #     $props = $action.ClientSideComponentProperties | ConvertFrom-Json
    #     $props | Format-List
    # } else {
    #     Write-Host "   (No properties defined)"
    # }
}

# Write-Host "`n📦 Application:"
# Get-PnPApp -Scope Tenant

# Connect-SPOService -Url "https://s63fb-admin.sharepoint.com"
# Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell

# $cred = Get-Credential
# Connect-SPOService -Url "https://s63fb-admin.sharepoint.com" -Credential $cred
# Set-SPOSite -Identity "https://s63fb.sharepoint.com/sites/selfserviceagent" -DenyAddAndCustomizePages $true
# Get-SPOSite -Identity "https://s63fb.sharepoint.com/sites/selfserviceagent" | Select DenyAddAndCustomizePages


# Step 5: Remove the app (optional)
# $appId = "c237d299-ef5b-42b5-9198-9f4e73095f9b"
# Remove-PnPCustomAction -Identity $appId -Scope Web -Force
# Write-Host "`n🧹 Removed Application Customizer with ID $appId"
#
#
#
# "features": [
#   {
#     "title": "service-desk-chat Feature",
#     "description": "The feature that activates elements of the service-desk-chat solution.",
#     "id": "2c91e7ff-e44c-4e76-890b-aa83fcff6fa2",
#     "version": "1.0.0.0",
#     "assets": {
#       "elementManifests": ["elements.xml", "ClientSideInstance.xml"]
#     }
#   }
# ]
# [
#   {
#     "AadAppId": "00000000-0000-0000-0000-000000000000",
#     "AadPermissions": null,
#     "AppCatalogVersion": "1.0.0.0",
#     "CanUpgrade": false,
#     "CDNLocation": "SharePoint Online",
#     "ContainsTenantWideExtension": false,
#     "CurrentVersionDeployed": true,
#     "Deployed": true,
#     "ErrorMessage": "No errors.",
#     "ID": "ddfc4e88-2eb0-444b-9030-c80020221fd9",
#     "InstalledVersion": "",
#     "IsClientSideSolution": true,
#     "IsEnabled": true,
#     "IsPackageDefaultSkipFeatureDeployment": true,
#     "IsValidAppPackage": true,
#     "ProductId": "105f66d9-0260-4c5d-819c-fb7e70a9b213",
#     "ShortDescription": "ServiceDeskChat description",
#     "SkipDeploymentFeature": false,
#     "StoreAssetId": "",
#     "SupportsTeamsTabs": true,
#     "ThumbnailUrl": "",
#     "Title": "sample-spfx-solution"
#   }
# ]
