using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)


Write-Host "Message Trace function triggered."
 
try {
    # -----------------------------
    # Read input from request body
    # -----------------------------
    $senderEmail    = $Request.Body.senderEmail
    $recipientEmail = $Request.Body.recipientEmail
    $subject        = $Request.Body.subject
    $days           = [int]$Request.Body.days
    $traceDate      = $Request.Body.traceDate
    
    # -----------------------------
    # Calculate date range
    # -----------------------------
    if ($traceDate) {
        try {
            $parsedDate = (Get-Date $traceDate).Date
            $startDate = $parsedDate
            $finalEndDate = $parsedDate.AddDays(1).AddSeconds(-1)
            # Write-Host "Searching for specific date: $($parsedDate.ToString('yyyy-MM-dd'))"
        }
        catch {
            throw "Invalid 'traceDate' format. Please provide a valid date (e.g., '2024-01-15')."
        }
    }
    else {
        if (-not $days -or $days -le 0) {
            throw "Parameter 'days' must be a positive number when 'traceDate' is not provided."
        }
        $finalEndDate = Get-Date
        $startDate = $finalEndDate.AddDays(-$days)
        # Write-Host "Searching for last $days day(s)"
    }
 
    # Used by pagination loop
    $currentEndDate = $finalEndDate
 
    # -----------------------------
    # Load cert-based auth variables
    # -----------------------------
    $CertBytes = [Convert]::FromBase64String($env:EXO_CERT_BASE64)
    $CertPassword = ConvertTo-SecureString $env:EXO_CERT_PASSWORD -AsPlainText -Force
 
    $Cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertBytes,
        $CertPassword,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable `
        -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
    )
 
    # -----------------------------
    # Connect to Exchange Online
    # -----------------------------
    Import-Module ExchangeOnlineManagement
 
    Connect-ExchangeOnline `
        -AppId $env:EXO_APP_ID `
        -Organization $env:EXO_TENANT_DNS `
        -Certificate $Cert `
        -ShowBanner:$false
 
    # -----------------------------
    # Message Trace pagination
    # -----------------------------
    $allResults = @()
    $resultSize = 5000
 
    do {
        $results = Get-MessageTraceV2 `
            -StartDate $startDate `
            -EndDate $currentEndDate `
            -SenderAddress $senderEmail `
            -RecipientAddress $recipientEmail `
            -ResultSize $resultSize `
            -ErrorAction Stop
 
        if ($subject) {
            $results = $results | Where-Object {
                $_.Subject -like "*$subject*"
            }
        }
 
        if ($results.Count -gt 0) {
            $allResults += $results
 
            # Get oldest message timestamp for next iteration
            $oldestMessageTime = ($results |
                Sort-Object Received |
                Select-Object -First 1).Received
 
            # Move end date backward to avoid duplicates
            $currentEndDate = $oldestMessageTime.AddSeconds(-1)
        }
 
    } while ($results.Count -eq $resultSize -and $currentEndDate -gt $startDate)
 
    # -----------------------------
    # Post-processing & Reporting
    # -----------------------------
    if ($allResults.Count -eq 0) {
        $summaryMessage = "Email not Found."
        $statusSummary = @()
        $mostRecentMessageDetails = $null
        $mostRecentMessageInfo = $null
       
        # Write-Host $summaryMessage
    }
    else {
        $summaryMessage = "$($allResults.Count) Email(s) found"
        # Write-Host $summaryMessage
 
        # -----------------------------
        # Status breakdown
        # -----------------------------
        $statusSummary = $allResults |
            Group-Object Status |
            ForEach-Object {
                [PSCustomObject]@{
                    Status = $_.Name
                    Count  = $_.Count
                }
            }
 
        # Log status breakdown
        # Write-Host "`nStatus Breakdown:"
        # foreach ($status in $statusSummary) {
        #     Write-Host "  - $($status.Status): $($status.Count)"
        # }
 
        # -----------------------------
        # Most recent message
        # -----------------------------
        $mostRecentMessage = $allResults |
            Sort-Object Received -Descending |
            Select-Object -First 1
 
        $mostRecentMessageInfo = @{
            MessageTraceId = $mostRecentMessage.MessageTraceId
            SenderAddress  = $mostRecentMessage.SenderAddress
            Recipient      = $mostRecentMessage.RecipientAddress
            Subject        = $mostRecentMessage.Subject
            Status         = $mostRecentMessage.Status
            Received       = $mostRecentMessage.Received
        }
 
        # Write-Host "`nMost Recent Message:"
        # Write-Host "  MessageTraceId: $($mostRecentMessage.MessageTraceId)"
        # Write-Host "  Recipient: $($mostRecentMessage.RecipientAddress)"
        # Write-Host "  Status: $($mostRecentMessage.Status)"
        # Write-Host "  Received: $($mostRecentMessage.Received)"
 
        # Fetch detailed trace
        try {
            # Write-Host "`nFetching detailed trace..."
            $mostRecentMessageDetails = Get-MessageTraceDetailV2 `
                -MessageTraceId $mostRecentMessage.MessageTraceId `
                -RecipientAddress $mostRecentMessage.RecipientAddress `
                -ErrorAction Stop
           
            # Write-Host "  Detail records retrieved: $($mostRecentMessageDetails.Count)"
        }
        catch {
            # Write-Warning "Unable to retrieve message trace details: $_"
            $mostRecentMessageDetails = @{
                error = "Unable to retrieve message trace details: $($_.Exception.Message)"
            }
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false
 
    # -----------------------------
    # Prepare response
    # -----------------------------
    $responseBody = @{
        summaryMessage          = $summaryMessage
        totalEmails             = $allResults.Count
        startDate               = $startDate.ToString("yyyy-MM-ddTHH:mm:ss")
        endDate                 = $finalEndDate.ToString("yyyy-MM-ddTHH:mm:ss")
        statusReport            = $statusSummary
        mostRecentMessage       = $mostRecentMessageInfo
        mostRecentMessageDetail = $mostRecentMessageDetails
    }
 
    # Write-Host ($responseBody | ConvertTo-Json -Depth 6)
    Write-Host "Message Trace function completed successfully."
 
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = ($responseBody | ConvertTo-Json -Depth 6)
        Headers    = @{
            "Content-Type" = "application/json"
        }
    })
}
catch {
    Write-Error "Error in Message Trace function: $_"
 
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body       = @{
            error = $_.Exception.Message
        } | ConvertTo-Json
    })
}