# Email-Size-Analysis-Filter-Script
This repository contains a PowerShell script to analyze and filter emails by size. It connects to Exchange Online, retrieves emails from the last 10 days, and categorizes them into internal, sent externally, and received externally. Results are exported to CSV files for analysis.

# The Script

This PowerShell script is designed to analyze and filter emails based on their size. It connects to Exchange Online, retrieves emails sent within the last 10 days, and categorizes them into internal emails, emails sent externally, and emails received from external sources. He also filter by message size, above 10Mo, 20Mo and 30Mo. The results are then exported to CSV files for further analysis.

## Prerequisites

- PowerShell
- ExchangeOnlineManagement module

## Installation

1. Install the ExchangeOnlineManagement module if not already installed:
    ```powershell
    Install-Module -Name ExchangeOnlineManagement
    ```

2. Clone this repository to your local machine.

## Customization
Make sure to customize the script by replacing `@domain.com` with your actual domain name to accurately filter internal and external emails.

## Customizing Paths
To ensure the script saves the output files to the correct locations on your system, you need to customize the file paths. You MUST change these paths to any directory of your choice.

## Usage

1. Open PowerShell and navigate to the directory where the script is located.

2. Run the script:
    ```powershell
    .\EmailSizeAnalysisScript.ps1
    ```

3. The script will connect to Exchange Online and retrieve emails sent within the last 10 days. It will filter the emails based on their size and categorize them into:
    - Internal emails
    - Emails sent externally
    - Emails received from external sources

4. The results will be exported to the following CSV files:
    - `internal-emails.csv`
    - `external-emails-sent.csv`
    - `external-emails-received.csv`

## Script Details

- **Connect to Exchange Online**:
    ```powershell
    Connect-ExchangeOnline
    ```

- **Define the dates for the last 10 days period**:
    ```powershell
    $startDate = (Get-Date).AddDays(-10).ToString("yyyy-MM-dd")
    $endDate = (Get-Date).ToString("yyyy-MM-dd")
    ```

- **Get emails sent with specific sizes**:
    ```powershell
    $results = Get-MessageTrace -StartDate $startDate -EndDate $endDate -PageSize 5000 |
    Where-Object {
       $_.Size -ge 10485760 -or  # Emails larger than 10 MB (10 MB = 10485760 bytes)
       $_.Size -ge 20971520 -or  # Emails larger than 20 MB (20 MB = 20971520 bytes)
       $_.Size -ge 31457280      # Emails larger than 30 MB (30 MB = 31457280 bytes)
    }
    ```

- **Filter to get internal emails**:
    ```powershell
    $internalEmails = $results | Where-Object { $_.SenderAddress -like "*@domain.com" -and $_.RecipientAddress -like "*@domain.com" }
    ```

- **Filter to get emails sent externally**:
    ```powershell
    $externalEmailsSent = $results | Where-Object { $_.SenderAddress -like "*@domain.com" -and $_.RecipientAddress -notlike "*@domain.com" }
    ```

- **Filter to get emails received from external sources**:
    ```powershell
    $externalEmailsReceived = $results | Where-Object { $_.SenderAddress -notlike "*@domain.com" -and $_.RecipientAddress -like "*@domain.com" }
    ```

- **Export internal emails**:
    ```powershell
    $internalEmails | Export-Csv -Path "C:\Scripts\emailsize\internal-emails.csv" -NoTypeInformation
    ```

- **Export emails sent externally**:
    ```powershell
    $externalEmailsSent | Export-Csv -Path "C:\Scripts\emailsize\external-emails-sent.csv" -NoTypeInformation
    ```

- **Export emails received from external sources**:
    ```powershell
    $externalEmailsReceived | Export-Csv -Path "C:\Scripts\emailsize\external-emails-received.csv" -NoTypeInformation
    ```

## Author

Script written by **David Bouhadana**.

- Blog: [M365 journey](https://m365journey.blog/)

## License

This project is licensed under **GNU GPL 3**. You are free to use, modify, and distribute this code as long as the modifications and derived versions are also licensed under GNU GPL 3. For more information, please refer to the full license text [GNU GPL 3](https://www.gnu.org/licenses/gpl-3.0.html).
