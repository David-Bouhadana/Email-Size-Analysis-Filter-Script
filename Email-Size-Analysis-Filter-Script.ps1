# Install the ExchangeOnlineManagement module if necessary
# Install-Module -Name ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Define the dates for the last 10 days period
$startDate = (Get-Date).AddDays(-10).ToString("yyyy-MM-dd")
$endDate = (Get-Date).ToString("yyyy-MM-dd")

# Get emails sent with specific sizes
$results = Get-MessageTrace -StartDate $startDate -EndDate $endDate -PageSize 5000 |
Where-Object {
   $_.Size -ge 10485760 -or  # Emails larger than 10 MB (10 MB = 10485760 bytes)
   $_.Size -ge 20971520 -or  # Emails larger than 20 MB (20 MB = 20971520 bytes)
   $_.Size -ge 31457280      # Emails larger than 30 MB (30 MB = 31457280 bytes)
}

# Filter to get internal emails
$internalEmails = $results | Where-Object { $_.SenderAddress -like "*@domain.com" -and $_.RecipientAddress -like "*@domain.com" }

# Filter to get emails sent externally
$externalEmailsSent = $results | Where-Object { $_.SenderAddress -like "*@domain.com" -and $_.RecipientAddress -notlike "*@domain.com" }

# Filter to get emails received from external sources
$externalEmailsReceived = $results | Where-Object { $_.SenderAddress -notlike "*@domain.com" -and $_.RecipientAddress -like "*@domain.com" }

# Export internal emails
$internalEmails | Export-Csv -Path "YourPath\internal-emails.csv" -NoTypeInformation

# Export emails sent externally
$externalEmailsSent | Export-Csv -Path "YourPath\external-emails-sent.csv" -NoTypeInformation

# Export emails received from external sources
$externalEmailsReceived | Export-Csv -Path "YourPath\external-emails-received.csv" -NoTypeInformation