$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'Platform'; Availability = 99.98; Incidents = 1; Status = 'Healthy' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Availability = 99.72; Incidents = 4; Status = 'Watch' }
    [pscustomobject]@{ Service = 'Remote access'; Owner = 'Security'; Availability = 98.84; Incidents = 7; Status = 'Action' }
)

New-OfficeMarkdown -Path '.\Service-Status.md' {
    MarkdownHeading -Level 1 -Text 'Service Status'
    MarkdownParagraph -Text 'The same PowerShell objects also feed the Word, Excel, PowerPoint, and PDF outputs in this recipe.'
    MarkdownTable -InputObject $services
    MarkdownHeading -Level 2 -Text 'Next steps'
    MarkdownTaskList -Items 'Review Remote access', 'Assign the incident action', 'Publish the approved status pack'
}

New-OfficeWord -Path '.\Service-Status.docx' {
    WordSection {
        WordHeader { WordParagraph -Text 'Weekly service status' -Style Heading2 }
        WordFooter { WordPageNumber -IncludeTotal }
        WordParagraph -Text 'Service Status' -Style Heading1
        WordParagraph -Text 'A document for owners who need narrative, tables, and an approval-ready file.'
        WordTable -InputObject $services -Style GridTable4Accent1 -Layout AutoFitToWindow {
            WordTableCondition -FilterScript { $_.Status -eq 'Action' } -BackgroundColor '#FEE2E2'
            WordTableCondition -FilterScript { $_.Status -eq 'Watch' } -BackgroundColor '#FEF3C7'
        }
        WordChart -Type Bar -InputObject $services -Category Service -Value Incidents -Title 'Incidents by service' -FitToPageWidth
    }
}

New-OfficeExcel -Path '.\Service-Status.xlsx' {
    ExcelSheet 'Services' {
        # Keep the chart's category and one numeric measure adjacent; owners and percentages remain in the table.
        ExcelTable -InputObject ($services | Select-Object Service,Incidents,Owner,Availability,Status) -Name 'ServiceStatus' -Address A1 -TableStyle 'TableStyleMedium9'
        ExcelSheetView -FreezeRows 1
        ExcelValidation -Type List -Table 'ServiceStatus' -Header Status -Values Healthy,Watch,Action
        ExcelConditionalFormat -Range 'D2:D4' -StartColor '#FEE2E2' -EndColor '#DCFCE7'
        ExcelChart -Range 'A1:B4' -Address A7 -Type ColumnClustered -Title 'Incidents by service' -Width 700px -Height 320px
        ExcelPageSetup -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 1
    }
    ExcelTableOfContents -Name 'Index' -BackLink -BackLinkText 'Back to Index'
}

New-OfficePowerPoint -Path '.\Service-Status.pptx' {
    Set-OfficePowerPointSlideSize -Preset Screen16x9
    PowerPointSlide -Title 'Weekly Service Status' -Subtitle 'Three services, one owner conversation' {
        PowerPointNotes 'Lead with the Remote access action and confirm its owner.'
    }
    PowerPointSlide -Title 'Current status' {
        PowerPointTable -InputObject $services
        PowerPointParagraph 'Decision: assign the Remote access action before the next service review.'
        PowerPointNotes 'Use the table for exact values and confirm the action owner.'
    }
    $focus = PowerPointSlide -Title 'Where to focus the review' -PassThru
    PowerPointChart -Slide $focus -InputObject $services -Category Service -Value Incidents -Type Column -Title 'Incidents by service' -X 90 -Y 135 -Width 780 -Height 290
    Add-OfficePowerPointTextBox -Slide $focus -Content { PowerPointText 'Availability remains a percentage in the status table; this chart compares incident counts.' -Size 18 -Color '#475569' } -X 60 -Y 450 -Width 840 -Height 55
    PowerPointNotes -Slide $focus -Text 'Discuss incident reduction separately from availability targets.'
}

New-OfficePdf -Path '.\Service-Status.pdf' -Theme Modern {
    PdfMetadata -Title 'Weekly service status' -Author 'Operations'
    PdfPageSetup -PageSize A4 -Margin 42
    PdfHeader 'Weekly service status'
    PdfFooter 'Page {page}/{pages}'
    PdfHeading 'Service Status' -Level 1
    PdfPanel 'Remote access requires an owner action before the next review.'
    PdfTable -InputObject $services -Property Service,Owner,Availability,Incidents,Status -HeaderBackgroundColor '#334155' -HeaderColor '#FFFFFF' -AutoFit
    PdfHeading 'Next steps' -Level 2
    PdfList -Items 'Review Remote access', 'Assign the incident action', 'Publish the approved status pack' -Numbered
}
