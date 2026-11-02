<#
.SYNOPSIS
Creates a readable service review workbook and an optional PDF copy.
.DESCRIPTION
Writes a titled, themed table with a percentage column and wrapped notes, then
fits the sheet to one page wide. PDF export reads the saved workbook and does
not require Microsoft Excel.
.PARAMETER OutputDirectory
Folder for Service-Review.xlsx and, when requested, Service-Review.pdf.
.PARAMETER IncludePdf
Also export the saved workbook to PDF.
.EXAMPLE
./Recipe-Excel-ReadableReview.ps1 -OutputDirectory ./Output/Review -IncludePdf
.OUTPUTS
System.IO.FileInfo for the workbook, and an OfficeConversionReport for the PDF.
#>
param(
    [string] $OutputDirectory = './ServiceReview',
    [switch] $IncludePdf
)

$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$workbook = Join-Path $OutputDirectory 'Service-Review.xlsx'
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

New-OfficeExcel -Path $workbook -Force -PassThru {
    ExcelSheet 'Review' {
        ExcelCell A1 'Service review'
        ExcelRow 1 -Bold -Height 26pt

        # The table header row (3) is frozen and the columns fitted by default.
        ExcelTable $services -Name Services -Address A3 -ColumnFormat @{ Availability = '0.00%' }
        ExcelColumn C -Width 54
        foreach ($row in 4..6) { ExcelRow $row -WrapText -AutoFit }

        # A PowerShell number stays numeric, so Excel can calculate with it.
        ExcelCell A8 'Availability target'
        ExcelCell B8 0.999 -NumberFormat '0.00%'
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
    }
}

if ($IncludePdf) {
    Convert-OfficeDocument -Path $workbook -OutputPath (Join-Path $OutputDirectory 'Service-Review.pdf')
}
