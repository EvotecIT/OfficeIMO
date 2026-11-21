using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a QR code or barcode to PDF content as a vector drawing.</summary>
/// <para>Two-dimensional codes default to 1.25 inches wide; linear barcodes to 0.33 mm per bar module, with their text underneath. The alternative text defaults to the symbology and the encoded text and is written to the tagged PDF.</para>
/// <example><summary>Add a QR code under an invoice summary.</summary>
/// <code>New-OfficePdf -Path .\Invoice.pdf {
///     PdfParagraph 'Pay online:'
///     PdfBarcode 'https://pay.example.com/inv/2026-0042' -Alignment Right
/// }</code></example>
/// <example><summary>Print a GS1-128 shipping label code on an existing document.</summary>
/// <code>PdfBarcode '(01)09501101530003(17)261231' -Type Gs1Code128 -Document $pdf -Width 8cm -PassThru</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePdfBarcode", DefaultParameterSetName = "Content")]
[Alias("PdfBarcode")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfBarcodeCommand : OfficeBarcodeCommandBase {
    /// <summary>Horizontal alignment in the content flow; left when omitted.</summary>
    [Parameter]
    public OfficeAlignment? Alignment { get; set; }

    /// <summary>Document to modify outside the PDF DSL.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true)]
    public PdfDocument? Document { get; set; }

    /// <summary>Return the modified document; only with Document.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var type = Type;
        string text = Text!;
        var result = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, "Document",
            content => content.AddBarcode(type, text, Configure, Alignment));
        if (PassThru && result != null) WriteObject(result);
    }
}
