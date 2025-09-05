param (
    [Parameter(Mandatory = $true)]
    [string]$siteUrl,

    [Parameter(Mandatory = $true)]
    [string]$clientId,

    [Parameter(Mandatory = $true)]
    [string]$tenant,

    [Parameter(Mandatory = $true)]
    [string]$botUrl,

    [Parameter(Mandatory = $true)]
    [string]$customScope,

    [Parameter(Mandatory = $true)]
    [string]$authority,

    [Parameter(Mandatory = $true)]
    [string]$certificatePath,

    [Parameter(Mandatory = $true)]
    [string]$certificatePassword,

    [Parameter(Mandatory = $false)]
    [string]$botName = "Service Desk Agent",

    [Parameter(Mandatory = $false)]
    [string]$buttonLabel = "Chat now",

    [Parameter(Mandatory = $false)]
    [string]$userEmail = "",

    [Parameter(Mandatory = $false)]
    [string]$botAvatarImage = "",

    [Parameter(Mandatory = $false)]
    [string]$botAvatarInitials = "SD"

)

Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificateBase64Encoded $certificatePath `
  -CertificatePassword (ConvertTo-SecureString $certificatePassword -AsPlainText -Force)
# Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenant -CertificatePath $certificatePath `
#   -CertificatePassword (ConvertTo-SecureString $certificatePassword -AsPlainText -Force)

$componentId = "6884b891-ab15-4bc6-b67e-b45a233cdd5f"

$propertiesObject = @{
  botURL = $botUrl
  botName = $botName
  buttonLabel = $buttonLabel
  userEmail = ""
  botAvatarImage = $botAvatarImage
  botAvatarInitials = $botAvatarInitials
  greet = $true
  customScope = $customScope
  clientID = $clientId
  authority = $authority
}
$properties = $propertiesObject | ConvertTo-Json -Compress

try {
    $existing = Get-PnPCustomAction -Scope Site | Where-Object { $_.Name -eq "BubbleChatCustomizer" }

    if ($existing) {
        Remove-PnPCustomAction -Identity $existing.Id -Scope Site -Force
        Write-Host "Removed existing BubbleChatCustomizer from site: $siteUrl"
    }

    Add-PnPCustomAction `
        -Name "BubbleChatCustomizer" `
        -Title "BubbleChat" `
        -Location "ClientSideExtension.ApplicationCustomizer" `
        -ClientSideComponentId $componentId `
        -ClientSideComponentProperties $properties `
        -Scope Site

    Write-Host "BubbleChat Application Customizer has been re-added to the site: $siteUrl"

} catch {
    Write-Host "Error during custom action update: $($_.Exception.Message)"
}
