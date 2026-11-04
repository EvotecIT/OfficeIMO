$path = '.\Project-Tracker.xlsx'
$tasks = @(
    [pscustomobject]@{ Task = 'Confirm scope'; Owner = 'Product'; Status = 'Done'; Progress = 100; Due = '2026-08-21' }
    [pscustomobject]@{ Task = 'Build API'; Owner = 'Engineering'; Status = 'In progress'; Progress = 70; Due = '2026-08-28' }
    [pscustomobject]@{ Task = 'Prepare pilot'; Owner = 'Operations'; Status = 'Blocked'; Progress = 25; Due = '2026-09-04' }
    [pscustomobject]@{ Task = 'Publish runbook'; Owner = 'Support'; Status = 'Not started'; Progress = 0; Due = '2026-09-08' }
)

New-OfficeExcel -Path $path {
    ExcelSheet 'Tracker' {
        ExcelTable -InputObject $tasks -Name 'ProjectTasks' -Address A1 -TableStyle 'TableStyleMedium9'
        ExcelSheetView -FreezeRows 1
        ExcelValidation -Type List -Table 'ProjectTasks' -Header 'Status' `
            -Values 'Not started', 'In progress', 'Blocked', 'Done'
        ExcelConditionalFormat -Range 'D2:D20' -StartColor '#FEE2E2' -EndColor '#DCFCE7'
        ExcelConditionalFormat -Table 'ProjectTasks' -Header 'Status' -Rule ContainsText -Text 'Blocked'
        ExcelChart -Range 'A1:D5' -Address A7 -Type BarClustered -Title 'Task progress' -Width 720px -Height 320px
        ExcelHeader -Center 'Project tracker'
        ExcelFooter -Right 'Page &P of &N'
        ExcelPageSetup -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
    }

    ExcelSheet 'Instructions' {
        ExcelCell -Address A1 -Value 'How to use this tracker'
        ExcelCell -Address A3 -Value '1. Add tasks to the ProjectTasks table.'
        ExcelCell -Address A4 -Value '2. Choose a status from the validation list.'
        ExcelCell -Address A5 -Value '3. Update progress; the color scale and chart follow the data.'
    }

    ExcelTableOfContents -Name 'Index' -BackLink -BackLinkText 'Back to Index'
}
