using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Recognizes an image as editable text, reading order and detected tables, with review evidence.</summary>
/// <para>Uses the original image resolution and the shared OfficeIMO layout engine. Animated and multi-page images are rejected.</para>
/// <example>
/// <summary>Recognize a ledger once and export its reviewed table.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>$image = Get-OfficeImageDocument -Path .\Ledger.png
/// $image | ConvertFrom-OfficeImage -OutputPath .\Ledger.xlsx -FirstRowIsHeader</code>
/// <para>Inspect $image.Ocr.Document.Tables and $image.Ocr.Pages before confirming the first row as headers.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeImageDocument")]
[OutputType(typeof(PdfSearchableOcrReview))]
public sealed class GetOfficeImageDocumentCommand : OfficeOcrCmdlet
{
    /// <summary>Path to a supported single-frame raster image.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional engine-neutral OCR provider. When omitted, uses the configured local Tesseract session.</summary>
    [Parameter]
    public IOcrEngine? Provider { get; set; }

    /// <summary>Advanced layout, confidence, orientation, scan preparation and resource limits.</summary>
    [Parameter]
    public PdfOcrMergeOptions? RecognitionOptions { get; set; }

    /// <summary>Minimum normalized confidence accepted into the editable document.</summary>
    [Parameter]
    [ValidateRange(0.0, 1.0)]
    public double? MinimumConfidence { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        string inputPath = PdfCommandUtilities.ResolveExistingFilePath(this, Path);
        var options = RecognitionOptions?.Clone() ?? new PdfOcrMergeOptions { ReconstructLayout = true };
        if (MinimumConfidence.HasValue) options.MinimumConfidence = MinimumConfidence.Value;
        CancelToken.ThrowIfCancellationRequested();
        var image = PdfImageDocumentSource.FromFile(inputPath, options.MaxRenderedBytesPerPage);
        IOcrEngine engine = Provider ?? (await TesseractOcr.CreateSessionAsync(CreateSessionOptions(), CancelToken)
            .ConfigureAwait(false)).Engine;
        var review = await image.PrepareSearchableOcrAsync(engine, options, CancelToken).ConfigureAwait(false);
        WriteObject(review);
    }
}
