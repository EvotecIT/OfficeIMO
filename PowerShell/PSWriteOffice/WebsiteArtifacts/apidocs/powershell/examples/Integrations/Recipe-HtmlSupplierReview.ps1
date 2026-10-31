<#
.SYNOPSIS
Turns an HTML supplier table into an editable Word review, PDF and Excel workbook.
.DESCRIPTION
PSParseHTML selects table id="offers". PSWriteOffice creates the documents.
The bundled input is synthetic. Prices use invariant decimal notation without group
separators and USD per unit. Up to 15 significant digits and 28 fractional places
are accepted, ignoring leading zeros and fractional padding zeros. Sub-cent prices
remain visible in both reports.
.EXAMPLE
./Recipe-HtmlSupplierReview.ps1 -OutputDirectory ./Output/SupplierReview
.EXAMPLE
./Recipe-HtmlSupplierReview.ps1 -HtmlPath ./MyOffers.html -OutputDirectory ./Output/MyReview
#>
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $HtmlPath
)

$ErrorActionPreference = 'Stop'
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $PSScriptRoot 'Output/SupplierReview' }
if (-not $HtmlPath) { $HtmlPath = Join-Path $PSScriptRoot 'Input/SupplierOffers.html' }
Import-Module PSParseHTML -MinimumVersion 3.0.3 -ErrorAction Stop
Import-Module PSWriteOffice -MinimumVersion 3.0.8 -ErrorAction Stop

$priceTextFormat = '0.00##########################'
$tables = @(ConvertFrom-HtmlTable -Path $HtmlPath -TableId offers -IncludeMetadata)
if ($tables.Count -ne 1 -or @($tables[0].Data).Count -eq 0) {
    throw 'Supply one nonempty HTML table with id="offers".'
}
$offers = @($tables[0].Data | ForEach-Object {
    if ([string]::IsNullOrWhiteSpace($_.Item) -or [string]::IsNullOrWhiteSpace($_.Supplier)) {
        throw 'Each offer needs Item and Supplier values.'
    }
    [decimal] $price = 0
    if (-not [decimal]::TryParse($_.UnitPriceUsd, [Globalization.NumberStyles]::AllowDecimalPoint, [Globalization.CultureInfo]::InvariantCulture, [ref] $price)) {
        throw 'UnitPriceUsd must be a nonnegative dot-decimal price without group separators.'
    }
    # Inspect the input, not the parsed value: Decimal.TryParse can round tiny quotes to zero.
    $priceParts = ([string] $_.UnitPriceUsd).Split('.')
    $fraction = if ($priceParts.Count -eq 2) { $priceParts[1].TrimEnd('0') } else { '' }
    $significantDigits = ($priceParts[0].TrimStart('0') + $fraction).TrimStart('0')
    $fractionalPlaces = $fraction.Length
    if ($significantDigits.Length -gt 15 -or $fractionalPlaces -gt 28) {
        throw 'UnitPriceUsd supports at most 15 significant digits and 28 fractional places, ignoring leading zeros and fractional padding zeros, so Word and numeric Excel quotes agree.'
    }
    $days = [int]::Parse($_.LeadDays, [Globalization.CultureInfo]::InvariantCulture)
    if ($price -lt 0 -or $days -lt 0) { throw 'Prices and lead times must be nonnegative.' }
    [pscustomobject]@{ Item = [string] $_.Item; Supplier = [string] $_.Supplier; UnitPriceUsd = $price; LeadDays = $days }
})
$reviewRows = $offers | Select-Object Item, Supplier,
    @{ Name = 'USD per unit'; Expression = { $_.UnitPriceUsd.ToString($priceTextFormat, [Globalization.CultureInfo]::InvariantCulture) } },
    @{ Name = 'Lead days'; Expression = { $_.LeadDays } }
$itemCount = @($offers | Select-Object -ExpandProperty Item -Unique).Count

$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$OutputDirectory = (Get-Item -LiteralPath $OutputDirectory).FullName
$docx = Join-Path $OutputDirectory 'SupplierReview.docx'
$pdf = Join-Path $OutputDirectory 'SupplierReview.pdf'
$xlsx = Join-Path $OutputDirectory 'SupplierOffers.xlsx'

New-OfficeWord -Path $docx {
    WordHeader -Type Default { WordParagraph -Text 'PROCUREMENT / OFFER REVIEW' }
    WordFooter -Type Default { WordPageNumber -IncludeTotal }
    WordParagraph {
        WordParagraphStyle -SpacingAfterPoints 12 -KeepWithNext $true
        WordText 'Supplier offer review' -Bold -FontSize 26 -Color '17365D'
    }
    WordParagraph {
        WordParagraphStyle -SpacingAfterPoints 10
        WordText "$($offers.Count) offers / $itemCount items" -Bold -FontSize 16 -Color '166534'
    }
    WordParagraph -Text 'Compare price and delivery together. The workbook retains numeric values for your own calculations.'
    WordParagraph {
        WordParagraphStyle -SpacingBeforePoints 16 -SpacingAfterPoints 8 -KeepWithNext $true
        WordText 'Offers received' -Bold -FontSize 15 -Color '17365D'
    }
    $table = WordTable -InputObject $reviewRows -Style GridTable4Accent1 -Layout AutoFitToWindow -PassThru
    $table | Get-OfficeWordTableCell | ForEach-Object {
        $_.MarginTop = [OfficeIMO.OfficeLength]::Centimeters(0.14); $_.MarginBottom = [OfficeIMO.OfficeLength]::Centimeters(0.14)
        $_.MarginLeft = [OfficeIMO.OfficeLength]::Centimeters(0.16); $_.MarginRight = [OfficeIMO.OfficeLength]::Centimeters(0.16)
    }
    WordParagraph {
        WordParagraphStyle -SpacingBeforePoints 16 -SpacingAfterPoints 8 -KeepWithNext $true
        WordText 'Before choosing an offer' -Bold -FontSize 15 -Color '17365D'
    }
    WordList -Style Bulleted {
        WordListItem -Text 'Confirm stock, quote expiry and the required delivery date.'
        WordListItem -Text 'Add freight and tax before comparing landed cost.'
        WordListItem -Text 'Record the purchasing decision and its owner in the editable document.'
    }
    WordParagraph {
        WordParagraphStyle -SpacingBeforePoints 12
        WordText 'Input contract: Item, Supplier, UnitPriceUsd and LeadDays. The bundled HTML contains synthetic data.' -FontSize 10 -Color '475569'
    }
}
Export-OfficeDocumentPdf -InputPath $docx -Path $pdf

New-OfficeExcel -Path $xlsx {
    ExcelSheet 'Offers' {
        ExcelGridlines -Hide
        ExcelCell -Address A1 -Value 'Supplier offers / USD per unit'
        ExcelRow -Row 1 -Height 30 -Bold $true -FontName 'Carlito'
        ExcelCell -Address A2 -Value 'Editable numeric values; freight and tax excluded.'
        ExcelTable -InputObject $offers -StartRow 4 -Name SupplierOffers -TableStyle TableStyleMedium2 -AutoFit
        for ($row = 5; $row -lt 5 + $offers.Count; $row++) {
            # Keep each quote's decimal places visible rather than rounding every price to cents.
            $priceText = $offers[$row - 5].UnitPriceUsd.ToString($priceTextFormat, [Globalization.CultureInfo]::InvariantCulture)
            $decimalPlaces = $priceText.Split('.')[1].Length
            ExcelCell -Row $row -Column 3 -Value $offers[$row - 5].UnitPriceUsd -NumberFormat ('0.' + ('0' * $decimalPlaces))
            ExcelRow -Row $row -Height 24
        }
        ExcelColumn -Column 1 -Width 25
        ExcelColumn -Column 2 -Width 22
        ExcelColumn -Column 3 -AutoFit
        ExcelColumn -Column 4 -Width 16
        ExcelFreeze -TopRows 4
    }
}
Get-Item -LiteralPath $docx, $pdf, $xlsx | Select-Object Name, Length
