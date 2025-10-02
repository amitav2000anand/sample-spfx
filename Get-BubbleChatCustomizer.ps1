# Define your connection parameters
$siteUrl = "https://s63fb.sharepoint.com/sites/selfserviceagent"
$clientId = "4ffd1a7a-9a30-48a9-bc3d-51060e46591b"
$tenant = "s63fb.onmicrosoft.com"
$certificatePassword = "YourStrongPassword123!"

# Load base64 certificate from file
$certBase64 = Get-Content -Path "cert_base64.txt" -Raw

# Connect to SharePoint
Connect-PnPOnline -Url $siteUrl `
  -ClientId $clientId `
  -Tenant $tenant `
  -CertificateBase64Encoded $certBase64 `
  -CertificatePassword (ConvertTo-SecureString $certificatePassword -AsPlainText -Force)

Write-Host "`n✅ Connected to SharePoint"

# Check for existing custom action
$existing = Get-PnPCustomAction -Scope Site | Where-Object { $_.Name -eq "BubbleChatCustomizer" }

if ($existing) {
    Write-Host "`n🔍 Found existing BubbleChatCustomizer:"
    $existing | Format-List Name, Id, Location, ClientSideComponentId, ClientSideComponentProperties
} else {
    Write-Host "`n❌ No BubbleChatCustomizer found on site: $siteUrl"
}
