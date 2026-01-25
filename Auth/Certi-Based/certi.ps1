$subject = "CN=EXO-AppOnly-MessageTrace"
$password = "Provide_Strong_Pass"
$folderPath = "C:\Azure\AD-App\Certi\"
$fileName = "EXO-AppOnly-MessageTrace"
$yearsValid = 2
 
New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
 
$cert = New-SelfSignedCertificate `
    -Subject $subject `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyUsage DigitalSignature `
    -NotAfter (Get-Date).AddYears($yearsValid)
 
$certPath = "Cert:\CurrentUser\My\$($cert.Thumbprint)"
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
 
Export-Certificate `
    -Cert $certPath `
    -FilePath "$folderPath\$fileName.cer"
 
Export-PfxCertificate `
    -Cert $certPath `
    -FilePath "$folderPath\$fileName.pfx" `
    -Password $securePassword