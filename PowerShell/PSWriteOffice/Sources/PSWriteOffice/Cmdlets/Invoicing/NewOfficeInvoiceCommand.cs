using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Creates an editable invoice without serializing or declaring standards compliance.</summary>
/// <example>
///   <summary>Creates an editable invoice without serializing or declaring standards compliance.</summary>
///   <code>$invoice = New-OfficeInvoice -Number INV-001 -IssueDate 2026-10-20 -Seller $seller -Buyer $buyer -Currency EUR</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeInvoice")]
[OutputType(typeof(Invoice))]
public sealed class NewOfficeInvoiceCommand : PSCmdlet {
    /// <summary>Seller-assigned document number.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Number { get; set; } = string.Empty;

    /// <summary>Explicit issue date.</summary>
    [Parameter(Mandatory = true)]
    public DateTime IssueDate { get; set; }

    /// <summary>Seller party, created with New-OfficeInvoiceParty or a property hashtable.</summary>
    [Parameter(Mandatory = true)]
    public InvoiceParty Seller { get; set; } = null!;

    /// <summary>Buyer party, created with New-OfficeInvoiceParty or a property hashtable.</summary>
    [Parameter(Mandatory = true)]
    public InvoiceParty Buyer { get; set; } = null!;

    /// <summary>ISO currency code.</summary>
    [Parameter]
    public string Currency { get; set; } = "EUR";

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(new Invoice { Number = Number, IssueDate = IssueDate, Currency = Currency, Seller = Seller, Buyer = Buyer });
    }
}
