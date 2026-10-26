$recipients = @(
    @{ FirstName = 'Ada'; OrderId = 'SO-1042'; DeliveryDate = '2026-09-01' }
    @{ FirstName = 'Grace'; OrderId = 'SO-1043'; DeliveryDate = '2026-09-03' }
)

foreach ($recipient in $recipients) {
    New-OfficeWord -Path ".\Order-$($recipient.OrderId).docx" -Force {
        WordHeading 'Order confirmation'
        WordParagraph {
            WordText 'Hello '
            WordField MergeField -Parameters '"FirstName"'
            WordText ','
        }
        WordParagraph {
            WordText 'Order '
            WordField MergeField -Parameters '"OrderId"'
            WordText ' is scheduled for '
            WordField MergeField -Parameters '"DeliveryDate"'
            WordText '.'
        }
        Invoke-OfficeWordMailMerge $recipient
    }
}
