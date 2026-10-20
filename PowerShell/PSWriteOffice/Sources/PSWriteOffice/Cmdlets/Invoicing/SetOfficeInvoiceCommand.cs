using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Updates common invoice header and party fields.</summary>
/// <example>
///   <summary>Updates common invoice header and party fields.</summary>
///   <code>$invoice | Set-OfficeInvoice -Number INV-002 -Currency EUR</code>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeInvoice")]
[OutputType(typeof(Invoice))]
public sealed class SetOfficeInvoiceCommand : PSCmdlet {
    /// <summary>Editable invoice.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public Invoice Invoice { get; set; } = null!;

    /// <summary>Replacement document number.</summary>
    [Parameter]
    public string? Number { get; set; }

    /// <summary>Replacement issue date.</summary>
    [Parameter]
    public DateTime? IssueDate { get; set; }

    /// <summary>Replacement currency code.</summary>
    [Parameter]
    public string? Currency { get; set; }

    /// <summary>Accounting VAT currency. Changing it invalidates its previously declared amount.</summary>
    [Parameter]
    public string? TaxCurrency { get; set; }

    /// <summary>Explicit VAT amount in the accounting currency; no exchange rate is inferred.</summary>
    [Parameter]
    public decimal? TaxAmountInAccountingCurrency { get; set; }

    /// <summary>Replacement seller.</summary>
    [Parameter]
    public InvoiceParty? Seller { get; set; }

    /// <summary>Replacement buyer.</summary>
    [Parameter]
    public InvoiceParty? Buyer { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Remove source-declared totals and tax breakdowns so the owner recomputes them from the edited model.</summary>
    [Parameter]
    public SwitchParameter RecalculateTotals { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Number != null) Invoice.Number = Number;
        if (IssueDate.HasValue) Invoice.IssueDate = IssueDate.Value;
        if (Currency != null) Invoice.Currency = Currency;
        if (Seller != null) Invoice.Seller = Seller;
        if (Buyer != null) Invoice.Buyer = Buyer;
        if (RecalculateTotals.IsPresent) {
            Invoice.ClearDeclaredAmounts();
        }
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TaxCurrency))) Invoice.TaxCurrency = TaxCurrency;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TaxAmountInAccountingCurrency))) Invoice.TaxAmountInAccountingCurrency = TaxAmountInAccountingCurrency;
        if (PassThru.IsPresent) WriteObject(Invoice);
    }
}
