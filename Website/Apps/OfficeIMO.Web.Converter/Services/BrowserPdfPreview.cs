using System.Globalization;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Pdf;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>Renders one bounded output page through the same PDF engine used by browser tools.</summary>
internal sealed class BrowserPdfPreview {
    private readonly PdfDocument _document;

    internal BrowserPdfPreview(byte[] bytes) {
        ArgumentNullException.ThrowIfNull(bytes);
        PdfLoadOptions options = BrowserPdfPolicy.CreateReadOptions(maximumInputBytes: BrowserPdfPolicy.MaxOutputBytes);
        _document = PdfDocument.Load(bytes, options);
        PageCount = _document.InspectForViewing().Pages.Count;
        if (PageCount == 0) throw new InvalidDataException("The output has no pages to preview.");
    }

    internal int PageCount { get; }

    internal PdfPageRenderResult Render(int pageNumber, CancellationToken cancellationToken = default) {
        if (pageNumber < 1 || pageNumber > PageCount) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        return _document.Render.Pages(pageNumber.ToString(CultureInfo.InvariantCulture), new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Scale = 1.5,
            ThumbnailMaxDimension = 1280,
            MaxPages = 1,
            MaxPixelsPerPage = 2_000_000,
            MaxOutputBytesPerPage = 8 * 1024 * 1024,
            MaxTotalOutputBytes = 8 * 1024 * 1024,
            Fonts = BrowserPortablePdfProfile.CreateDrawingFonts(),
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
            RenderTimeout = TimeSpan.FromSeconds(15),
            ContinueOnError = true
        }, cancellationToken).Single();
    }
}
