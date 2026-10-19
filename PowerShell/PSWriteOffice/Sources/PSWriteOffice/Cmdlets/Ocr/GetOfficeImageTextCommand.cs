using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Recognizes text in an image with automatic local OCR runtime discovery.</summary>
/// <example>
///   <summary>Read English text from an image.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeImageText -Path .\Scan.png</code>
///   <para>Returns recognized text and automatically uses an installed Tesseract runtime.</para>
/// </example>
/// <example>
///   <summary>Read English and Polish text with recognition evidence.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeImageText -Path .\Scan.png -Language English, Polish -PassThru</code>
///   <para>Returns the full OCR result, including confidence, word geometry, provider, model, and diagnostics.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeImageText")]
[OutputType(typeof(string))]
[OutputType(typeof(OcrResult))]
public sealed class GetOfficeImageTextCommand : OfficeOcrCmdlet
{
    /// <summary>PNG, JPEG, TIFF, BMP, GIF, WebP, or JPEG 2000 image path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Return the complete OCR result instead of recognized text only.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        TesseractOcrSessionOptions sessionOptions = CreateSessionOptions();
        ValidateSupportedImagePath(Path);
        string inputPath = PdfCommandUtilities.ResolveExistingFilePath(this, Path);
        OcrResult result = await TesseractOcr
            .RecognizeFileAsync(inputPath, sessionOptions, CancelToken)
            .ConfigureAwait(false);
        WriteObject(PassThru.IsPresent ? result : result.Text);
    }

    private static void ValidateSupportedImagePath(string path)
    {
        switch (System.IO.Path.GetExtension(path).ToLowerInvariant())
        {
            case ".png":
            case ".jpg":
            case ".jpeg":
            case ".tif":
            case ".tiff":
            case ".bmp":
            case ".gif":
            case ".webp":
            case ".jp2":
            case ".j2k":
                return;
            default:
                throw new PSArgumentException(
                    "Tesseract OCR supports PNG, JPEG, TIFF, BMP, GIF, WebP, and JPEG 2000 image files.",
                    nameof(Path));
        }
    }
}
