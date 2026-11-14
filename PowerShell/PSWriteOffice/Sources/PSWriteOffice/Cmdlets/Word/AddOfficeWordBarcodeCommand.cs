using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Visuals;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a QR code or barcode to a Word paragraph.</summary>
/// <para>The code is an SVG picture with a PNG fallback, so Word draws it sharply and older readers still show it; PDF conversion keeps it as vector content. Two-dimensional codes default to 1.25 inches wide; linear barcodes to 0.33 mm per bar module, with their text underneath. The alternative text defaults to the symbology and the encoded text.</para>
/// <example><summary>Add a QR code that opens a web page.</summary>
/// <code>New-OfficeWord -Path .\Flyer.docx {
///     WordParagraph 'Scan for the event agenda:'
///     WordBarcode 'https://evotec.xyz/events'
/// }</code></example>
/// <example><summary>Print an order number as a Code 128 barcode.</summary>
/// <code>WordParagraph { WordText 'Order: '; WordBarcode 'ORD-2026-0042' -Type Code128 -Width 6cm }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeWordBarcode")]
[OutputType(typeof(WordImage))]
[Alias("WordBarcode")]
public sealed class AddOfficeWordBarcodeCommand : OfficeBarcodeCommandBase {
    /// <summary>Target paragraph; the current DSL paragraph, or a new one, is used when omitted.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Return the native editable WordImage.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordParagraph paragraph = Paragraph ?? ResolveParagraph();
        WordImage image = paragraph.AddBarcode(Type, Text!, Configure);
        if (PassThru) WriteObject(image);
    }

    private WordParagraph ResolveParagraph() {
        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
