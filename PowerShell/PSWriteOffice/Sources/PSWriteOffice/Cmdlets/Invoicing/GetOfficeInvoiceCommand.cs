using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;

namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Reads CII or UBL while retaining unmapped-source evidence for later saves.</summary>
/// <example>
///   <summary>Reads CII or UBL while retaining unmapped-source evidence for later saves.</summary>
///   <code>Get-Item ./invoice.xml | Get-OfficeInvoice</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeInvoice")]
[OutputType(typeof(Invoice))]
public sealed class GetOfficeInvoiceCommand : PSCmdlet {
    /// <summary>Literal invoice XML path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "FilePath", "LiteralPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Return the source declaration and unmapped-data diagnostics around the editable model.</summary>
    [Parameter]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var result = InvoiceCommandUtilities.Read(OfficeFileCommandUtilities.ResolvePath(this, Path));
        if (!result.HasCompleteMapping) WriteWarning("Some source business data is not mapped; rewriting is blocked unless that loss is explicitly accepted.");
        WriteObject(AsResult.IsPresent ? result : result.Invoice);
    }
}
