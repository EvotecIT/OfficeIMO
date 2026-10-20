using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Returns editable invoice lines, optionally selected by their stable identifier.</summary>
/// <example>
///   <summary>Returns editable invoice lines, optionally selected by their stable identifier.</summary>
///   <code>$invoice | Get-OfficeInvoiceLine</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeInvoiceLine")]
[OutputType(typeof(InvoiceLine))]
public sealed class GetOfficeInvoiceLineCommand : PSCmdlet {
    /// <summary>Editable invoice.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public Invoice Invoice { get; set; } = null!;
    /// <summary>Optional exact line identifiers.</summary>
    [Parameter]
    public string[]? Id { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var line in Invoice.Lines) if (Id == null || Id.Contains(line.Id, StringComparer.Ordinal)) WriteObject(line);
    }
}
