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

New-OfficePdf -Path $path -Theme Modern {
    PdfMetadata -Title 'Service invoice INV-2026-0818' -Author 'Northwind Automation'
    PdfPageSetup -PageSize A4 -Margin 42
    PdfHeader 'NORTHWIND AUTOMATION'
    PdfFooter 'Invoice INV-2026-0818 | Page {page}/{pages}'

    PdfHeading 'SERVICE INVOICE' -Level 1 -Color '#0F766E'
    PdfParagraph {
        PdfText 'Invoice: ' -Bold
        PdfText 'INV-2026-0818'
        PdfText '    Issue date: ' -Bold
        PdfText '2026-08-18'
    }
    PdfParagraph {
        PdfText 'Bill to: ' -Bold
        PdfText 'Contoso Ltd., 1 Example Street, London'
    }
    PdfHorizontalRule -Color '#0F766E' -SpacingBefore 12 -SpacingAfter 14

    PdfTable -InputObject $lines -Property Description, Quantity, UnitPrice, Amount `
        -Header 'Description', 'Qty', 'Unit price', 'Amount' `
        -HeaderBackgroundColor '#0F766E' -HeaderColor '#FFFFFF' `
        -AutoFit

    PdfSpacer -Height 14
    PdfText -Text ('Subtotal: ' + $subtotal.ToString('C2', $currency)) -Alignment Right -Bold
    PdfText -Text ('Illustrative tax (' + $taxRate.ToString('0.##%', $currency) + '): ' + $tax.ToString('C2', $currency)) -Alignment Right
    PdfText -Text ('Total due: ' + $total.ToString('C2', $currency)) -Alignment Right -Bold -Size 14 -Color '#0F766E'

    PdfPanel 'Payment terms: 14 days. Include the invoice number with the transfer.'
    PdfHeading 'Questions' -Level 2
    PdfParagraph {
        PdfText 'Contact '
        PdfText 'billing@example.com' -Link 'mailto:billing@example.com' -Color '#2563EB'
        PdfText ' before the due date if any line needs correction.'
    }
}
