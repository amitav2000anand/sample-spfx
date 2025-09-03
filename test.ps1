# Step 1: Create the certificate
$cert = New-SelfSignedCertificate `
  -Subject "CN=TestCert" `
  -CertStoreLocation "Cert:\CurrentUser\My" `
  -KeyExportPolicy Exportable `
  -KeySpec Signature `
  -KeyLength 2048 `
  -HashAlgorithm SHA256 `
  -NotAfter (Get-Date).AddYears(1)

# Step 2: Export the private key (.pfx) for GitHub CLI login
$pwd = ConvertTo-SecureString -String "YourStrongPassword123!" -Force -AsPlainText
Export-PfxCertificate `
  -Cert $cert `
  -FilePath ".\TestCert.pfx" `
  -Password $pwd

# ✅ Step 3: Export the public certificate (.cer) for Azure AD
Export-Certificate `
  -Cert $cert `
  -FilePath ".\TestCert.cer"
