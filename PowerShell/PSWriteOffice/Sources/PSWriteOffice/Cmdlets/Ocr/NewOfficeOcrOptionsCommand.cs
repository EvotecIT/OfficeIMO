using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Creates optional OCR settings without starting a provider or downloading language data.</summary>
/// <example><summary>Read a scanned image with selected languages.</summary>
/// <code>$ocr = New-OfficeOcrOptions -Language English,Polish -NoLanguageDownload
/// Read-OfficeDocument ./scan.png -As Text -OcrOptions $ocr</code></example>
[Cmdlet(VerbsCommon.New, "OfficeOcrOptions")]
[OutputType(typeof(OfficeOcrOptions))]
public sealed class NewOfficeOcrOptionsCommand : OfficeOcrCmdlet {
    /// <summary>Optional engine-neutral OCR provider.</summary>
    [Parameter] public IOcrEngine? Provider { get; set; }
    /// <summary>Bounded recognition and enrichment policy for document views.</summary>
    [Parameter] [OfficeOptionsTransformation(typeof(OfficeDocumentOcrExecutionOptions))]
    public OfficeDocumentOcrExecutionOptions? ExecutionOptions { get; set; }
    /// <summary>Layout, confidence and resource policy for original image or PDF reviews.</summary>
    [Parameter] [OfficeOptionsTransformation(typeof(PdfOcrMergeOptions))]
    public PdfOcrMergeOptions? RecognitionOptions { get; set; }
    /// <summary>Minimum confidence accepted into an editable OCR review.</summary>
    [Parameter] [ValidateRange(0D, 1D)] public double? MinimumConfidence { get; set; }
    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        var recognition = RecognitionOptions?.Clone() ?? new PdfOcrMergeOptions { ReconstructLayout = true };
        if (MinimumConfidence.HasValue) recognition.MinimumConfidence = MinimumConfidence.Value;
        WriteObject(new OfficeOcrOptions {
            Provider = Provider, SessionOptions = CreateSessionOptions(),
            ExecutionOptions = ExecutionOptions?.Clone() ?? new OfficeDocumentOcrExecutionOptions(), RecognitionOptions = recognition
        });
        return Task.CompletedTask;
    }
}
