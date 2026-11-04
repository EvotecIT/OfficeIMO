<#
.SYNOPSIS
Creates a readable service review workbook and an optional PDF copy.
.DESCRIPTION
Uses worksheet pipelines to write typed values, size columns, wrap notes, and fit row
heights without a DSL block. Saves the workbook once and always closes it. PDF export
uses the saved workbook and does not require Microsoft Excel.
.PARAMETER OutputDirectory
Folder for Service-Review.xlsx and, when requested, Service-Review.pdf.
.PARAMETER IncludePdf
Also export the saved workbook to PDF.
.EXAMPLE
./Recipe-Excel-ReadableReview.ps1 -OutputDirectory ./Output/Review -IncludePdf
.OUTPUTS
System.IO.FileInfo. The workbook and optional PDF files.
#>
param(
    [string] $OutputDirectory = (Join-Path (Get-Location) 'ServiceReview'),
    [switch] $IncludePdf
)

Import-Module PSWriteOffice -ErrorAction Stop
$outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$workbookPath = Join-Path $outputPath 'Service-Review.xlsx'
$services = @(
    [pscustomobject]@{
        Service = 'Identity'
        Availability = 0.9995
        Notes = 'Access review completed. Remove the remaining temporary assignments before the next review.'
    }
    [pscustomobject]@{
        Service = 'Messaging'
        Availability = 0.9982
        Notes = 'Mailbox migration is on track. Validate the shared mailbox permissions with the support team.'
    }
    [pscustomobject]@{
        Service = 'Backup'
        Availability = 1.0
        Notes = 'Restore exercise passed. Record the next recovery exercise in the operations calendar.'
    }
)

$workbook = New-OfficeExcel -Path $workbookPath
try {
    $sheet = $workbook | Add-OfficeExcelSheet -Name 'Review' -PassThru
    Set-OfficeExcelRow -Sheet $sheet -Row 1 -Values 'Service review' -Bold:$true -Height 26
    Add-OfficeExcelTable -Sheet $sheet -InputObject $services -Address A3 -Name 'Services'

    Set-OfficeExcelColumn -Sheet $sheet -Column A -Width 20
    Set-OfficeExcelColumn -Sheet $sheet -Column B -Width 16
    Set-OfficeExcelColumn -Sheet $sheet -Column C -Width 54
    for ($row = 4; $row -le 6; $row++) {
        Set-OfficeExcelRow -Sheet $sheet -Row $row -WrapText:$true -AutoFit
        Set-OfficeExcelCell -Sheet $sheet -Row $row -Column 2 -NumberFormat '0.00%'
    }

    # A PowerShell number remains numeric, so Excel can calculate with it.
    Set-OfficeExcelCell -Sheet $sheet -Address A8 -Value 'Availability target'
    Set-OfficeExcelCell -Sheet $sheet -Address B8 -Value 0.999 -NumberFormat '0.00%'
    $workbook | Set-OfficeExcelSheetView -Sheet 'Review' -FreezeRows 3
    $workbook | Set-OfficeExcelPageSetup -Sheet 'Review' -FitToWidth 1 -FitToHeight 0
    $workbook | Save-OfficeExcel
} finally {
    $workbook | Close-OfficeExcel
}

Get-Item -LiteralPath $workbookPath
if ($IncludePdf) {
    $pdfPath = Join-Path $outputPath 'Service-Review.pdf'
    Export-OfficeDocumentPdf -InputPath $workbookPath -Path $pdfPath
    Get-Item -LiteralPath $pdfPath
}
