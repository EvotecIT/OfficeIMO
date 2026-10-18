$path = '.\Project-Tracker.xlsx'
$tasks = @(
    [pscustomobject]@{ Task = 'Confirm scope'; Owner = 'Product'; Status = 'Done'; Progress = 100; Due = '2026-08-21' }
    [pscustomobject]@{ Task = 'Build API'; Owner = 'Engineering'; Status = 'In progress'; Progress = 70; Due = '2026-08-28' }
    [pscustomobject]@{ Task = 'Prepare pilot'; Owner = 'Operations'; Status = 'Blocked'; Progress = 25; Due = '2026-09-04' }
    [pscustomobject]@{ Task = 'Publish runbook'; Owner = 'Support'; Status = 'Not started'; Progress = 0; Due = '2026-09-08' }
)

ExcelNew -Path $path {
    ExcelSheet 'Tracker' {
        ExcelTable -Data $tasks -TableName 'ProjectTasks' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium9' -AutoFit
        ExcelFreeze -TopRows 1
        ExcelValidationList -TableName 'ProjectTasks' -HeaderName 'Status' `
            -Values 'Not started', 'In progress', 'Blocked', 'Done'
        ExcelConditionalColorScale -Range 'D2:D20' -StartColor '#FEE2E2' -EndColor '#DCFCE7'
        ExcelConditionalRule -TableName 'ProjectTasks' -HeaderName 'Status' -RuleType ContainsText -Text 'Blocked'
        ExcelChart -Range 'A1:D5' -Row 7 -Column 1 -Type BarClustered -Title 'Task progress' -WidthPixels 720 -HeightPixels 320
        ExcelHeaderFooter -HeaderCenter 'Project tracker' -FooterRight 'Page &P of &N'
        ExcelOrientation -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
    }

    ExcelSheet 'Instructions' {
        ExcelCell -Address A1 -Value 'How to use this tracker'
        ExcelCell -Address A3 -Value '1. Add tasks to the ProjectTasks table.'
        ExcelCell -Address A4 -Value '2. Choose a status from the validation list.'
        ExcelCell -Address A5 -Value '3. Update progress; the color scale and chart follow the data.'
    }

    ExcelTableOfContents -SheetName 'Index' -AddBackLinks -BackLinkText 'Back to Index'
}
