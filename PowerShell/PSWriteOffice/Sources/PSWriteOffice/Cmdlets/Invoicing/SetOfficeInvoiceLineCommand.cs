using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Edits selected line fields while retaining other invoice content.</summary>
/// <example>
///   <summary>Edits selected line fields while retaining other invoice content.</summary>
///   <code>$invoice | Get-OfficeInvoiceLine -Id 1 | Set-OfficeInvoiceLine -Quantity 3</code>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeInvoiceLine")]
[OutputType(typeof(InvoiceLine))]
public sealed class SetOfficeInvoiceLineCommand : PSCmdlet {
    /// <summary>Editable invoice line.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public InvoiceLine Line { get; set; } = null!;
    /// <summary>Description.</summary>
    [Parameter]
    public string? Name { get; set; }
    /// <summary>Quantity.</summary>
    [Parameter]
    public decimal? Quantity { get; set; }
    /// <summary>Net unit price.</summary>
    [Parameter]
    public decimal? UnitPrice { get; set; }
    /// <summary>VAT percentage.</summary>
    [Parameter]
    public decimal? TaxRate { get; set; }
    /// <summary>VAT category.</summary>
    [Parameter]
    public string? TaxCategory { get; set; }
    /// <summary>Return the changed line.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Name))) Line.Name = Name ?? string.Empty;
        if (Quantity.HasValue) Line.Quantity = Quantity.Value;
        if (UnitPrice.HasValue) Line.UnitPrice = UnitPrice.Value;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TaxRate))) Line.Tax.Rate = TaxRate;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TaxCategory))) Line.Tax.Code = TaxCategory ?? string.Empty;
        if (PassThru.IsPresent) WriteObject(Line);
    }
}
