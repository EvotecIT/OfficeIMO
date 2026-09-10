using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Source and prepared page appearances for explicit scan review.</summary>
public sealed class PdfScanPreview {
    private readonly byte[] _source, _processed;
    internal PdfScanPreview(int pageNumber, byte[] source, byte[] processed, double width, double height,
        int pixelWidth, int pixelHeight, IReadOnlyList<string> diagnostics, OfficeScanProcessingReport? report) {
        PageNumber = pageNumber; _source = (byte[])source.Clone(); _processed = (byte[])processed.Clone();
        Width = width; Height = height; PixelWidth = pixelWidth; PixelHeight = pixelHeight;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray()); ScanProcessing = report;
    }
    /// <summary>One-based source page.</summary>
    public int PageNumber { get; }
    /// <summary>Prepared width in points at the requested sampling density.</summary>
    public double Width { get; }
    /// <summary>Prepared height in points at the requested sampling density.</summary>
    public double Height { get; }
    /// <summary>Prepared width in pixels.</summary>
    public int PixelWidth { get; }
    /// <summary>Prepared height in pixels.</summary>
    public int PixelHeight { get; }
    /// <summary>Geometry and tone decisions from optional affine processing.</summary>
    public OfficeScanProcessingReport? ScanProcessing { get; }
    /// <summary>Rendering and preparation limits or decisions.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Independent original page PNG.</summary>
    public byte[] GetSourcePng() => (byte[])_source.Clone();
    /// <summary>Independent prepared page PNG.</summary>
    public byte[] GetPreparedPng() => (byte[])_processed.Clone();
    /// <summary>Creates a single-page raster PDF of the reviewed appearance. Interactive content and native text are not copied.</summary>
    public PdfDocument CreateImagePdf(CancellationToken cancellationToken = default) => PdfDocument.CreateFromImages(
        new[] { new PdfImageDocumentSource(_processed) }, new PdfImageDocumentOptions {
            FixedPageSize = new PageSize(Width, Height),
            Margin = 0,
            Fit = OfficeImageFit.Stretch
        }, cancellationToken);
}

internal static partial class PdfOcr {
    internal static async Task<PdfScanPreview> PreviewScanAsync(byte[] source, int pageNumber,
        PdfOcrMergeOptions options, PdfLoadOptions? loadOptions, CancellationToken token) {
        options.Validate(); token.ThrowIfCancellationRequested();
        var document = PdfReadDocument.Open(source, loadOptions, token);
        if (pageNumber < 1 || pageNumber > document.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        if (!options.GetSelectedPages(document.Pages.Count).Contains(pageNumber))
            throw new ArgumentException("The preview page does not belong to the selected OCR pages.", nameof(pageNumber));
        var rendered = PdfPageImageRenderer.RenderPage(document, pageNumber, new PdfPageRenderOptions {
            Dpi = options.Dpi,
            Format = PdfPageRenderFormat.Png,
            ImageCodec = options.ImageCodec,
            MaxPixelsPerPage = options.MaxPixelsPerPage,
            MaxOutputBytesPerPage = options.MaxRenderedBytesPerPage,
            ContinueOnError = false
        }, token);
        (double width, double height) = document.Pages[pageNumber - 1].GetInteractionPageSize();
        var request = new OcrRequest {
            Payload = rendered.Bytes!,
            MediaType = "image/png",
            PageNumber = pageNumber,
            PixelWidth = rendered.Width,
            PixelHeight = rendered.Height,
            RegionCoordinateUnit = OcrCoordinateUnit.Points,
            Region = new OcrRegion { Width = width, Height = height }
        };
        PreparedPage prepared = await PreparePageAsync(request, null, options, token).ConfigureAwait(false);
        return new PdfScanPreview(pageNumber, rendered.Bytes!, request.Payload, request.Region!.Width, request.Region.Height,
            request.PixelWidth!.Value, request.PixelHeight!.Value, rendered.Diagnostics.Concat(prepared.Diagnostics).ToArray(), prepared.Report);
    }
}