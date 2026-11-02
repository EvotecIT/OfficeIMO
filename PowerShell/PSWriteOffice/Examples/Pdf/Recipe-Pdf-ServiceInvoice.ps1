$path = '.\Service-Invoice.pdf'
$items = @(
    [pscustomobject]@{ Description = 'Automation assessment'; Quantity = 1; UnitPrice = [decimal]1200 }
    [pscustomobject]@{ Description = 'Report implementation'; Quantity = 3; UnitPrice = [decimal]850 }
    [pscustomobject]@{ Description = 'Handover workshop'; Quantity = 1; UnitPrice = [decimal]600 }
)
$currency = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$taxRate = [decimal]0.20 # Illustrative rate; use the rate and rounding rules appropriate to your invoice.
$subtotal = [decimal]0
$lines = foreach ($item in $items) {
    $amount = [decimal]::Round($item.Quantity * $item.UnitPrice, 2, [MidpointRounding]::AwayFromZero)
    $subtotal += $amount
    [pscustomobject]@{
        Description = $item.Description
        Quantity = $item.Quantity
        UnitPrice = $item.UnitPrice.ToString('C2', $currency)
        Amount = $amount.ToString('C2', $currency)
    }
}
$tax = [decimal]::Round($subtotal * $taxRate, 2, [MidpointRounding]::AwayFromZero)
$total = $subtotal + $tax

New-OfficePdf -Path $path {
    PdfTheme Report
    PdfMetadata -Title 'Service invoice INV-2026-0818' -Author 'Northwind Automation'
    PdfPageSetup -PageSize A4 -Margin 42
    PdfHeader 'NORTHWIND AUTOMATION'
    PdfFooter 'Invoice INV-2026-0818 | Page {page}/{pages}'

    PdfHeading 'SERVICE INVOICE' -Level 1 -Color '#0F766E'
    PdfText -Run @(
        @{ Text = 'Invoice: '; Bold = $true }
        @{ Text = 'INV-2026-0818' }
        @{ Text = '    Issue date: '; Bold = $true }
        @{ Text = '2026-08-18' }
    )
    PdfText -Run @(
        @{ Text = 'Bill to: '; Bold = $true }
        @{ Text = 'Contoso Ltd., 1 Example Street, London' }
    )
    PdfHr -Color '#0F766E' -SpacingBefore 12 -SpacingAfter 14

    PdfTable -InputObject $lines -Property Description, Quantity, UnitPrice, Amount `
        -Header 'Description', 'Qty', 'Unit price', 'Amount' `
        -HeaderFill '#0F766E' -HeaderTextColor '#FFFFFF' `
        -RightAlignNumeric -AutoFitColumns

    PdfSpacer -Height 14
    PdfText -Text ('Subtotal: ' + $subtotal.ToString('C2', $currency)) -Align Right -Bold
    PdfText -Text ('Illustrative tax (' + $taxRate.ToString('0.##%', $currency) + '): ' + $tax.ToString('C2', $currency)) -Align Right
    PdfText -Text ('Total due: ' + $total.ToString('C2', $currency)) -Align Right -Bold -FontSize 14 -Color '#0F766E'

    PdfPanel 'Payment terms: 14 days. Include the invoice number with the transfer.'
    PdfHeading 'Questions' -Level 2
    PdfText -Run @(
        @{ Text = 'Contact ' }
        @{ Text = 'billing@example.com'; LinkUri = 'mailto:billing@example.com'; Color = '#2563EB' }
        @{ Text = ' before the due date if any line needs correction.' }
    )
}
