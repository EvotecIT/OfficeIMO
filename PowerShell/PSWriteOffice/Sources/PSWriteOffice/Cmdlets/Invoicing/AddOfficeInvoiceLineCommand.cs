using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Adds an invoice line with explicit quantity, price and VAT.</summary>
/// <example>
///   <summary>Adds an invoice line with explicit quantity, price and VAT.</summary>
///   <code>$invoice | Add-OfficeInvoiceLine -Name Consulting -Quantity 1 -UnitPrice 100 -TaxRate 19</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeInvoiceLine")]
[OutputType(typeof(InvoiceLine))]
public sealed class AddOfficeInvoiceLineCommand : PSCmdlet {
    /// <summary>Editable invoice.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public Invoice Invoice { get; set; } = null!;

    /// <summary>Line description.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Quantity.</summary>
    [Parameter(Mandatory = true)]
    public decimal Quantity { get; set; }

    /// <summary>Net unit price.</summary>
    [Parameter(Mandatory = true)]
    public decimal UnitPrice { get; set; }

    /// <summary>VAT percentage; omission is retained for categories that prohibit a rate.</summary>
    [Parameter]
    public decimal? TaxRate { get; set; }

    /// <summary>UNCL 5305 VAT category, such as S, Z or E.</summary>
    [Parameter]
    public string TaxCategory { get; set; } = "S";

    /// <summary>Exemption reason when required by the category.</summary>
    [Parameter]
    public string? ExemptionReason { get; set; }

    /// <summary>Optional stable line identifier; omission uses the next line number.</summary>
    [Parameter]
    public string? Id { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var line = new InvoiceLine {
            Id = Id ?? string.Empty,
            Name = Name, Quantity = Quantity, UnitPrice = UnitPrice,
            Tax = new InvoiceTaxCategory { Code = TaxCategory, Rate = TaxRate, ExemptionReason = ExemptionReason }
        };
        Invoice.AddLine(line);
        if (PassThru.IsPresent) WriteObject(line);
    }
}
