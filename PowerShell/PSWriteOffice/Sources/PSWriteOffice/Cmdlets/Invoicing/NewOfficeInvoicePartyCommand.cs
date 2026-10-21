using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Creates a seller or buyer from ordinary party and address parameters.</summary>
/// <example>
///   <summary>Creates a seller or buyer from ordinary party and address parameters.</summary>
///   <code>$seller = New-OfficeInvoiceParty -Name Example -CountryCode DE -VatId DE123456789</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeInvoiceParty")]
[OutputType(typeof(InvoiceParty))]
public sealed class NewOfficeInvoicePartyCommand : PSCmdlet {
    /// <summary>Legal party name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>ISO two-letter country code.</summary>
    [Parameter(Mandatory = true)]
    public string CountryCode { get; set; } = string.Empty;

    /// <summary>VAT registration identifier.</summary>
    [Parameter]
    public string? VatId { get; set; }

    /// <summary>First address line.</summary>
    [Parameter]
    public string? AddressLine { get; set; }

    /// <summary>City.</summary>
    [Parameter]
    public string? City { get; set; }

    /// <summary>Postal code.</summary>
    [Parameter]
    public string? PostCode { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var party = new InvoiceParty {
            Name = Name, Address = new InvoiceAddress { CountryCode = CountryCode, Line1 = AddressLine, City = City, PostCode = PostCode }
        };
        if (VatId != null) party.TaxRegistrations.Add(new InvoiceTaxRegistration(VatId, InvoiceTaxRegistration.VatScheme));
        WriteObject(party);
    }
}
