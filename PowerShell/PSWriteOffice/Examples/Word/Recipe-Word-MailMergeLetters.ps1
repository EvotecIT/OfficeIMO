$recipients = @(
    @{ FirstName = 'Ada'; OrderId = 'SO-1042'; DeliveryDate = '2026-09-01' }
    @{ FirstName = 'Grace'; OrderId = 'SO-1043'; DeliveryDate = '2026-09-03' }
)

foreach ($recipient in $recipients) {
    $path = ".\Order-$($recipient.OrderId).docx"

    WordNew -Path $path {
        WordSection {
            WordParagraph -Text 'Order confirmation' -Style Heading1
            WordParagraph {
                WordText 'Hello '
                WordField -Type MergeField -Parameters '"FirstName"'
                WordText ','
            }
            WordParagraph {
                WordText 'Order '
                WordField -Type MergeField -Parameters '"OrderId"'
                WordText ' is scheduled for '
                WordField -Type MergeField -Parameters '"DeliveryDate"'
                WordText '.'
            }
            Invoke-OfficeWordMailMerge -Values $recipient
        }
    }
}
