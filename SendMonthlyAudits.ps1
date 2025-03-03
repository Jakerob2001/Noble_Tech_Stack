# Define the property mapping hash table
#     "PCL" = @("pclo365audit@noblehousehotels.com")
$propertyMapping = @{
    "ARG" = @("argo365audit@noblehousehotels.com")
    "CAB" = @("cabo365audit@noblehousehotels.com")
    "CHA" = @("chio365audit@noblehousehotels.com")
    "EDG" = @("edgo365audit@noblehousehotels.com")
    "ELJ" = @("eljo365audit@noblehousehotels.com")
    "GWC" = @("gwco365audit@noblehousehotels.com")
    "IOF" = @("iofo365audit@noblehousehotels.com")
    "JICR" = @("jicro365audit@noblehousehotels.com")
    "KKR" = @("kkro365audit@noblehousehotels.com")
    "LAP" = @("lapo365audit@noblehousehotels.com")
    "LDM" = @("ldmo365audit@noblehousehotels.com")
    "LPI" = @("lpio365audit@noblehousehotels.com")
    "MAR" = @("maro365audit@noblehousehotels.com")
    "MBR" = @("mbro365audit@noblehousehotels.com")
    "NHHR" = @("jake.robinson@noblehousehotels.com")
    "NVWT" = @("nvwto365audit@noblehousehotels.com")
    "OKR" = @("okro365audit@noblehousehotels.com")
    "CKA" = @("pclo365audit@noblehousehotels.com")
    "HCA" = @("pclo365audit@noblehousehotels.com")
    "HCL" = @("pclo365audit@noblehousehotels.com")
    "PBR" = @("pbro365audit@noblehousehotels.com")
    "POK" = @("jake.robinson@noblehousehotels.com")
    "POR" = @("poro365audit@noblehousehotels.com")
    "RTI" = @("rtio365audit@noblehousehotels.com")
    "SOL" = @("solo365audit@noblehousehotels.com")
    "SRC" = @("srco365audit@noblehousehotels.com")
    "TR" = @("tro365audit@noblehousehotels.com")
    "TSH" = @("tsho365audit@noblehousehotels.com")
}

foreach ($propertyCode in $propertyMapping.Keys) {

    # Define the email details
    $subject = "CORRECTED - $propertyCode January O365 Users"
    $body = @"
<html>
<body>
    <p>Hello,</p>
    <p>Here is a list of all active O365 accounts at <b>$propertyCode</b>. Please review and let us know of any users that can be deactivated.</p>
    <p>Best regards,<br>NHHR Tech Support</p>
</body>
</html>
"@


    $toRecipients = $propertyMapping[$propertyCode]
    $sender = "techsupport@noblehousehotels.com"
    $ccRecipients = @("jake.robinson@noblehousehotels.com", "gbrashear@noblehousehotels.com")

    # Create the list of To recipient objects
    $toRecipientObjects = @()
    foreach ($toRecipient in $toRecipients) {
        $toRecipientObjects += @{
            EmailAddress = @{
                Address = $toRecipient
            }
        }
    }

    # Create the list of CC recipient objects
    $ccRecipientObjects = @()
    foreach ($ccRecipient in $ccRecipients) {
        $ccRecipientObjects += @{
            EmailAddress = @{
                Address = $ccRecipient
            }
        }
    }

    $attachmentFilePath = "C:\Users\$env:USERNAME\OneDrive - Noble House Hotels & Resorts\Documents\Monthly Audits\2025 January Audits\January $propertyCode - O365 Users.xlsx"
    $attachmentName = [System.IO.Path]::GetFileName($attachmentFilePath)

    # Read the attachment file and convert it to a base64 string
    $attachmentContentBytes = [System.IO.File]::ReadAllBytes($attachmentFilePath)
    $attachmentContent = [System.Convert]::ToBase64String($attachmentContentBytes)

    # Create the message object with an attachment
    $params = @{
        Message = @{
            Subject = $subject
            Body = @{
                ContentType = "HTML"
                Content = $body
            }
            Attachments = @(
                @{
                    '@odata.type' = "#microsoft.graph.fileAttachment"
                    Name = $attachmentName
                    ContentBytes = $attachmentContent
                }
            )
            ToRecipients = $toRecipientObjects
            CcRecipients = $ccRecipientObjects
            From = @{
                EmailAddress = @{
                    Address = $sender
                }
            }
        }
        SaveToSentItems = $true
    }

    # Send the email
    Send-MgUserMail -UserId "jake.robinson@noblehousehotels.com" -BodyParameter $params
}