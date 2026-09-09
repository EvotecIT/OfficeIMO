using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Pdf;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>Renders one bounded output page through the same PDF engine used by browser tools.</summary>
internal sealed class BrowserPdfPreview {
    private readonly PdfDocument _document;
    private readonly bool _canExtractContent;
    private readonly Lazy<OfficeFontFaceCollection> _fonts = new(BrowserPortablePdfProfile.CreateDrawingFonts);

    internal BrowserPdfPreview(byte[] bytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(TimeSpan.FromSeconds(15));
        deadline.Token.ThrowIfCancellationRequested();
        PdfLoadOptions options = BrowserPdfPolicy.CreateReadOptions(maximumInputBytes: BrowserPdfPolicy.MaxOutputBytes);
        _document = PdfDocument.Load(bytes, options);
        PdfDocumentViewInfo geometry = _document.InspectForViewing(includeLogicalContent: false, cancellationToken: deadline.Token);
        PageCount = geometry.PageCount;
        _canExtractContent = geometry.CanExtractContent;
        if (PageCount == 0) throw new InvalidDataException("The output has no pages to preview.");
    }

    internal int PageCount { get; }

    internal PdfPageRenderResult Render(int pageNumber, CancellationToken cancellationToken = default) {
        if (pageNumber < 1 || pageNumber > PageCount) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        cancellationToken.ThrowIfCancellationRequested();
        if (_canExtractContent) return _document.Render.Pages(pageNumber.ToString(CultureInfo.InvariantCulture), new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Scale = 1.5,
            ThumbnailMaxDimension = 1280,
            MaxPages = 1,
            MaxPixelsPerPage = 2_000_000,
            MaxOutputBytesPerPage = 8 * 1024 * 1024,
            MaxTotalOutputBytes = 8 * 1024 * 1024,
            Fonts = _fonts.Value,
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
            RenderTimeout = TimeSpan.FromSeconds(15),
            ContinueOnError = true
        }, cancellationToken).Single();
        return _document.Render.DisplayPage(pageNumber, new PdfPageDisplayOptions {
            Scale = 1.5,
            MaximumDimension = 1280,
            MaximumPixels = 2_000_000,
            MaximumOutputBytes = 8 * 1024 * 1024,
            Timeout = TimeSpan.FromSeconds(15)
        }, cancellationToken);
    }
}
