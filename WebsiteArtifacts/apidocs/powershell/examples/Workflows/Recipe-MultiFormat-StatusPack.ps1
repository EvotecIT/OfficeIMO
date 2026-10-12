$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'Platform'; Availability = 99.98; Incidents = 1; Status = 'Healthy' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Availability = 99.72; Incidents = 4; Status = 'Watch' }
    [pscustomobject]@{ Service = 'Remote access'; Owner = 'Security'; Availability = 98.84; Incidents = 7; Status = 'Action' }
)

MarkdownNew -Path '.\Service-Status.md' {
    MarkdownHeading -Level 1 -Text 'Service Status'
    MarkdownParagraph -Text 'The same PowerShell objects also feed the Word, Excel, PowerPoint, and PDF outputs in this recipe.'
    MarkdownTable -InputObject $services
    MarkdownHeading -Level 2 -Text 'Next steps'
    MarkdownTaskList -Items 'Review Remote access', 'Assign the incident action', 'Publish the approved status pack'
}

WordNew -Path '.\Service-Status.docx' {
    WordSection {
        WordHeader { WordParagraph -Text 'Weekly service status' -Style Heading2 }
        WordFooter { WordPageNumber -IncludeTotalPages }
        WordParagraph -Text 'Service Status' -Style Heading1
        WordParagraph -Text 'A document for owners who need narrative, tables, and an approval-ready file.'
        WordTable -InputObject $services -Style GridTable4Accent1 -Layout AutoFitToWindow {
            WordTableCondition -FilterScript { $_.Status -eq 'Action' } -BackgroundColor '#FEE2E2'
            WordTableCondition -FilterScript { $_.Status -eq 'Watch' } -BackgroundColor '#FEF3C7'
        }
        WordChart -Type Bar -Data $services -CategoryProperty Service -SeriesProperty Incidents -Title 'Incidents by service' -FitToPageWidth
    }
}

ExcelNew -Path '.\Service-Status.xlsx' {
    ExcelSheet 'Services' {
        # Keep the chart's category and one numeric measure adjacent; owners and percentages remain in the table.
        ExcelTable -Data ($services | Select-Object Service,Incidents,Owner,Availability,Status) -TableName 'ServiceStatus' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium9' -AutoFit
        ExcelFreeze -TopRows 1
        ExcelValidationList -TableName 'ServiceStatus' -HeaderName Status -Values Healthy,Watch,Action
        ExcelConditionalColorScale -Range 'D2:D4' -StartColor '#FEE2E2' -EndColor '#DCFCE7'
        ExcelChart -Range 'A1:B4' -Row 7 -Column 1 -Type ColumnClustered -Title 'Incidents by service' -WidthPixels 700 -HeightPixels 320
        ExcelOrientation -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 1
    }
    ExcelTableOfContents -SheetName 'Index' -AddBackLinks -BackLinkText 'Back to Index'
}

PptNew -Path '.\Service-Status.pptx' {
    PptSlideSize -Preset Screen16x9
    PptSlide {
        PptTitle -Title 'Weekly Service Status'
        PptTextBox -Text 'Three services, one owner conversation' -X 90 -Y 190 -Width 700 -Height 70
        PptNotes -Text 'Lead with the Remote access action and confirm its owner.'
    }
    PptSlide {
        PptTitle -Title 'Current status'
        PptTable -Data $services -X 55 -Y 130 -Width 840 -Height 210
        PptTextBox -Run @{ Text = 'Decision: assign the Remote access action before the next service review.'; FontSize = 20; Color = '#0F172A' } -X 60 -Y 390 -Width 840 -Height 75
        PptNotes -Text 'Use the table for exact values and confirm the action owner.'
    }
    PptSlide {
        PptTitle -Title 'Where to focus the review'
        PptChart -Data $services -CategoryProperty Service -SeriesProperty Incidents -Type ClusteredColumn -Title 'Incidents by service' -X 90 -Y 135 -Width 780 -Height 290
        PptTextBox -Run @{ Text = 'Availability remains a percentage in the status table; this chart compares incident counts.'; FontSize = 18; Color = '#475569' } -X 60 -Y 450 -Width 840 -Height 55
        PptNotes -Text 'Discuss incident reduction separately from availability targets.'
    }
}

PdfNew -Path '.\Service-Status.pdf' {
    PdfTheme Report
    PdfMetadata -Title 'Weekly service status' -Author 'Operations'
    PdfPageSetup -PageSize A4 -Margin 42
    PdfHeader 'Weekly service status'
    PdfFooter 'Page {page}/{pages}'
    PdfHeading 'Service Status' -Level 1
    PdfPanel 'Remote access requires an owner action before the next review.'
    PdfTable -InputObject $services -Property Service,Owner,Availability,Incidents,Status -HeaderFill '#334155' -HeaderTextColor '#FFFFFF' -AutoFitColumns -RightAlignNumeric
    PdfHeading 'Next steps' -Level 2
    PdfList -Items 'Review Remote access', 'Assign the incident action', 'Publish the approved status pack' -Numbered
}
