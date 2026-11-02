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

New-OfficePdf -Path $path -Theme Modern -Force {
    PdfMetadata -Title 'Service invoice INV-2026-0818' -Author 'Northwind Automation'
    PdfPageSetup -PageSize A4 -Margin 15mm
    PdfHeader 'NORTHWIND AUTOMATION'
    PdfFooter { PdfText 'Invoice INV-2026-0818 | Page '; PdfPageNumber -IncludeTotal }

    PdfHeading 'Service invoice'
    PdfParagraph {
        PdfText 'Invoice: ' -Bold; PdfText 'INV-2026-0818'
        PdfText '    Issue date: ' -Bold; PdfText '2026-08-18'
    }
    PdfParagraph { PdfText 'Bill to: ' -Bold; PdfText 'Contoso Ltd., 1 Example Street, London' }
    PdfHorizontalRule -SpacingBefore 12pt -SpacingAfter 14pt

    PdfTable $lines -Header 'Description', 'Qty', 'Unit price', 'Amount' -ColumnAlignment Left, Right, Right, Right -AutoFit

    PdfSpacer 14pt
    PdfParagraph -Alignment Right { PdfText 'Subtotal: ' -Bold; PdfText $subtotal.ToString('C2', $currency) }
    PdfParagraph "Illustrative tax ($($taxRate.ToString('0.##%', $currency))): $($tax.ToString('C2', $currency))" -Alignment Right
    PdfParagraph -Alignment Right { PdfText "Total due: $($total.ToString('C2', $currency))" -Bold -Size 14pt -Color Accent1 }

    PdfPanel 'Payment terms: 14 days. Include the invoice number with the transfer.'
    PdfHeading 'Questions' -Level 2
    PdfParagraph {
        PdfText 'Contact '
        PdfText 'billing@example.com' -Link 'mailto:billing@example.com'
        PdfText ' before the due date if any line needs correction.'
    }
}
