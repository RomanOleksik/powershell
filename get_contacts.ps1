# 1. Connect to Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# 2. Get the Sent Items folder (5 represents the specific ID for Sent Items)
$SentItems = $Namespace.GetDefaultFolder(5)
$Items = $SentItems.Items

# 3. Create a list to hold unique emails
$EmailList = @()

Write-Host "Scanning Sent Items... This may take a moment."

# 4. Loop through the latest 1000 emails (Adjust limit as needed for performance)
# To scan ALL, remove the counter logic, but be patient.
$Counter = 0
foreach ($Item in $Items) {
    if ($Item.Class -eq 43) { # 43 = MailItem
        foreach ($Recipient in $Item.Recipients) {
            # Create a custom object for the data
            $Object = [PSCustomObject]@{
                Name  = $Recipient.Name
                Email = $Recipient.Address
            }
            $EmailList += $Object
        }
    }
    $Counter++
    if ($Counter -ge 1000) { break } 
}

# 5. Deduplicate and Export to CSV
$UniqueEmails = $EmailList | Select-Object -Unique Email, Name
$UniqueEmails | Export-Csv -Path "C:\Temp\OutlookSentContacts.csv" -NoTypeInformation

Write-Host "Export Complete. File saved to C:\Temp\OutlookSentContacts.csv"
