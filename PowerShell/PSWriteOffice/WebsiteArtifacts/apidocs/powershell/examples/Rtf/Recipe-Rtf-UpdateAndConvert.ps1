$source = '.\Service-Note.rtf'
$updated = '.\Service-Note-Updated.rtf'
$markdown = '.\Service-Note.md'

New-OfficeRtf -Path $source -Text 'Service: Identity', 'Status: Draft'
Update-OfficeRtfText -Path $source -OutputPath $updated -OldText 'Draft' -NewText 'Approved' -AppendParagraph 'Reviewed by Platform'
ConvertFrom-OfficeRtf -Path $updated -As Markdown -OutputPath $markdown
