$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'Platform'; Availability = 99.98; Incidents = 1; Status = 'Healthy' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Availability = 99.72; Incidents = 4; Status = 'Watch' }
    [pscustomobject]@{ Service = 'Remote access'; Owner = 'Security'; Availability = 98.84; Incidents = 7; Status = 'Action' }
)
$nextSteps = 'Review Remote access', 'Assign the incident action', 'Publish the approved status pack'
$theme = 'Modern'   # One theme gives Word, Excel, PowerPoint and PDF the same fonts, colors and table style.

New-OfficeMarkdown -Path .\Service-Status.md -Force {
    MarkdownHeading 'Service status'
    MarkdownParagraph 'The same PowerShell objects also feed the Word, Excel, PowerPoint and PDF outputs in this recipe.'
    MarkdownTable $services
    MarkdownHeading 'Next steps' -Level 2
    MarkdownTaskList $nextSteps
}

New-OfficeWord -Path .\Service-Status.docx -Theme $theme -Force {
    WordHeader { WordParagraph 'Weekly service status' }
    WordFooter { WordPageNumber -IncludeTotal }
    WordHeading 'Service status'
    WordParagraph 'A document for owners who need narrative, tables and an approval-ready file.'
    WordTable $services {
        WordTableCondition { $_.Status -eq 'Action' } -BackgroundColor Danger
        WordTableCondition { $_.Status -eq 'Watch' } -BackgroundColor Warning
    }
    WordChart -Type Bar -InputObject $services -Category Service -Value Incidents -Title 'Incidents by service'
}

New-OfficeExcel -Path .\Service-Status.xlsx -Theme $theme -Force {
    ExcelSheet 'Services' {
        # Service and Incidents come first so the chart range stays two adjacent columns.
        ExcelTable ($services | Select-Object Service, Incidents, Owner, Availability, Status) -Name ServiceStatus
        ExcelValidation -Type List -Table ServiceStatus -Header Status -Values Healthy, Watch, Action
        ExcelConditionalFormat -Table ServiceStatus -Header Availability -StartColor Danger -EndColor Success
        ExcelChart 'A1:B4' -Address A7 -Type Column -Title 'Incidents by service'
        ExcelPageSetup -Orientation Landscape -FitToWidth 1 -FitToHeight 1
    }
    ExcelTableOfContents -Name Index -BackLink -BackLinkText 'Back to Index'
}

New-OfficePowerPoint -Path .\Service-Status.pptx -Theme $theme -Force {
    PptSlide -Title 'Weekly service status' -Subtitle 'Three services, one owner conversation' {
        PptNotes 'Lead with the Remote access action and confirm its owner.'
    }
    PptSlide -Title 'Current status' -Subtitle 'Decision: assign the Remote access action before the next service review.' {
        PptTable $services
        PptNotes 'Use the table for exact values and confirm the action owner.'
    }
    PptSlide -Title 'Where to focus the review' -Subtitle 'Availability stays in the status table; this chart compares incident counts.' {
        PptChart -InputObject $services -Category Service -Value Incidents -Type Column -Title 'Incidents by service'
        PptNotes 'Discuss incident reduction separately from availability targets.'
    }
}

New-OfficePdf -Path .\Service-Status.pdf -Theme $theme -Force {
    PdfMetadata -Title 'Weekly service status' -Author 'Operations'
    PdfPageSetup -PageSize A4 -Margin 15mm
    PdfHeader 'Weekly service status'
    PdfFooter { PdfText 'Page '; PdfPageNumber -IncludeTotal }
    PdfHeading 'Service status'
    PdfPanel 'Remote access requires an owner action before the next review.'
    PdfTable $services
    PdfHeading 'Next steps' -Level 2
    PdfList -Numbered { foreach ($step in $nextSteps) { PdfListItem $step } }
}
