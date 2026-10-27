using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Invoicing;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Invoicing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Invoicing.Pdf;
namespace PSWriteOffice.Cmdlets.Invoicing;

/// <summary>Creates a visible PDF and embedded invoice XML from the same captured model.</summary>
/// <example>
///   <summary>Creates a visible PDF and embedded invoice XML from the same captured model.</summary>
///   <code>$invoice | Export-OfficeInvoicePdf -Path ./invoice.pdf</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeInvoicePdf", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeInvoicePdfCommand : PSCmdlet {
    /// <summary>Editable invoice.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public Invoice Invoice { get; set; } = null!;

    /// <summary>Destination PDF filename.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Exact supported CII specification for the embedded invoice.</summary>
    [Parameter]
    public InvoiceSpecificationRelease SpecificationRelease { get; set; } = InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2;

    /// <summary>Embedded invoice profile.</summary>
    [Parameter]
    public InvoiceProfile Profile { get; set; } = InvoiceProfile.En16931;

    /// <summary>Optional existing PDF rendering options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfCreateOptions? PdfOptions { get; set; }

    /// <summary>Replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path, "Export invoice PDF")) return;
        var target = new InvoiceXmlOptions(SpecificationRelease, InvoiceSyntax.Cii, Profile);
        var sourceBytes = InvoiceCommandUtilities.Write(Invoice, target, allowUnmappedLoss: false);
        var captured = InvoiceParser.Read(sourceBytes);
        var snapshot = PdfInvoiceDocument.Create(captured.Invoice, target);
        OfficeFileOutput.Write(path, snapshot.ToPdfBytes(PdfOptions), Force.IsPresent);
        if (PassThru.IsPresent) WriteObject(new FileInfo(path));
    }
}
