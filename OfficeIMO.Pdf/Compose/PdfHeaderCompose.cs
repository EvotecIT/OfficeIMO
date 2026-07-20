using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Header builder (alignment, text, page number tokens).</summary>
public class PdfHeaderCompose {
    private readonly PdfOptions _opts;
    internal PdfHeaderCompose(PdfOptions opts) { _opts = opts; }
    /// <summary>Sets header alignment to the left.</summary>
    public PdfHeaderCompose AlignLeft() { _opts.HeaderAlign = PdfAlign.Left; return this; }
    /// <summary>Sets header alignment to the center.</summary>
    public PdfHeaderCompose AlignCenter() { _opts.HeaderAlign = PdfAlign.Center; return this; }
    /// <summary>Sets header alignment to the right.</summary>
    public PdfHeaderCompose AlignRight() { _opts.HeaderAlign = PdfAlign.Right; return this; }
    /// <summary>Sets header text color.</summary>
    public PdfHeaderCompose Color(PdfColor color) { _opts.HeaderTextColor = color; return this; }
    /// <summary>Sets header font.</summary>
    public PdfHeaderCompose Font(PdfStandardFont font) { _opts.HeaderFont = font; _opts.HeaderFontFamily = null; return this; }
    /// <summary>Uses a registered named font family for header text.</summary>
    public PdfHeaderCompose FontFamily(string familyName) { _opts.HeaderFontFamily = familyName; return this; }
    /// <summary>Sets header font size in points.</summary>
    public PdfHeaderCompose FontSize(double size) { Guard.Positive(size, nameof(size)); _opts.HeaderFontSize = size; return this; }
    /// <summary>Sets header baseline offset above the top margin in points.</summary>
    public PdfHeaderCompose Offset(double points) { Guard.NonNegative(points, nameof(points)); _opts.HeaderOffsetY = points; return this; }
    /// <summary>Renders the current page number in the header.</summary>
    public PdfHeaderCompose PageNumber() { _opts.ClearHeaderSegmentsForCompose(); _opts.ClearHeaderZonesForCompose(); _opts.ShowHeader = true; _opts.HeaderFormat = "{page}"; return this; }
    /// <summary>Renders the current page number and total pages in the header.</summary>
    public PdfHeaderCompose PageNumberWithTotal() { _opts.ClearHeaderSegmentsForCompose(); _opts.ClearHeaderZonesForCompose(); _opts.ShowHeader = true; _opts.HeaderFormat = "{page}/{pages}"; return this; }
    /// <summary>Renders left, center, and right header zones on one line. Supports {page} and {pages}.</summary>
    public PdfHeaderCompose Zones(string? left, string? center, string? right) {
        _opts.SetHeaderZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders page-1-only left, center, and right header zones. Supports {page} and {pages}.</summary>
    public PdfHeaderCompose FirstPageZones(string? left, string? center, string? right) {
        _opts.SetFirstPageHeaderZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders even-page-only left, center, and right header zones. Supports {page} and {pages}.</summary>
    public PdfHeaderCompose EvenPagesZones(string? left, string? center, string? right) {
        _opts.SetEvenPageHeaderZonesForCompose(left, center, right);
        return this;
    }

    /// <summary>Adds an image to the running header.</summary>
    public PdfHeaderCompose Image(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) =>
        Image(data, width, height, align, fit, alternativeText: null);

    /// <summary>Adds a meaningful image to the running header with alternate text.</summary>
    public PdfHeaderCompose Image(byte[] data, double width, double height, string? alternativeText) =>
        Image(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the running header.</summary>
    public PdfHeaderCompose Image(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        _opts.AddHeaderImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds an image to the page-1-only header.</summary>
    public PdfHeaderCompose FirstPageImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) =>
        FirstPageImage(data, width, height, align, fit, alternativeText: null);

    /// <summary>Adds a meaningful image to the page-1-only header with alternate text.</summary>
    public PdfHeaderCompose FirstPageImage(byte[] data, double width, double height, string? alternativeText) =>
        FirstPageImage(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the page-1-only header.</summary>
    public PdfHeaderCompose FirstPageImage(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        _opts.AddFirstPageHeaderImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds an image to the even-page-only header.</summary>
    public PdfHeaderCompose EvenPagesImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) =>
        EvenPagesImage(data, width, height, align, fit, alternativeText: null);

    /// <summary>Adds a meaningful image to the even-page-only header with alternate text.</summary>
    public PdfHeaderCompose EvenPagesImage(byte[] data, double width, double height, string? alternativeText) =>
        EvenPagesImage(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the even-page-only header.</summary>
    public PdfHeaderCompose EvenPagesImage(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        _opts.AddEvenPageHeaderImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds a shape to the running header.</summary>
    public PdfHeaderCompose Shape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddHeaderShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Adds a shape to the page-1-only header.</summary>
    public PdfHeaderCompose FirstPageShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddFirstPageHeaderShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Adds a shape to the even-page-only header.</summary>
    public PdfHeaderCompose EvenPagesShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddEvenPageHeaderShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Renders a literal header text format. Supports {page} and {pages}.</summary>
    public PdfHeaderCompose Text(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearHeaderSegmentsForCompose();
        _opts.ClearHeaderZonesForCompose();
        _opts.ShowHeader = true;
        _opts.HeaderFormat = format;
        return this;
    }

    /// <summary>Builds a custom header from text and page tokens.</summary>
    /// <param name="build">Delegate to compose header segments.</param>
    public PdfHeaderCompose Text(System.Action<HeaderTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetHeaderSegmentsForCompose();
        _opts.ClearHeaderZonesForCompose();
        var builder = new HeaderTextBuilder(segments);
        build(builder);
        return this;
    }

    /// <summary>Renders a page-1-only header text format. Supports {page} and {pages}.</summary>
    public PdfHeaderCompose FirstPageText(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearFirstPageHeaderSegmentsForCompose();
        _opts.ClearFirstPageHeaderZonesForCompose();
        _opts.DifferentFirstPageHeaderFooter = true;
        _opts.FirstPageHeaderFormat = format;
        return this;
    }

    /// <summary>Builds a page-1-only header from text and page tokens.</summary>
    /// <param name="build">Delegate to compose first-page header segments.</param>
    public PdfHeaderCompose FirstPageText(System.Action<HeaderTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetFirstPageHeaderSegmentsForCompose();
        _opts.ClearFirstPageHeaderZonesForCompose();
        var builder = new HeaderTextBuilder(segments);
        build(builder);
        return this;
    }

    /// <summary>Renders an even-page-only header text format. Supports {page} and {pages}.</summary>
    public PdfHeaderCompose EvenPagesText(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearEvenPageHeaderSegmentsForCompose();
        _opts.ClearEvenPageHeaderZonesForCompose();
        _opts.DifferentOddAndEvenPagesHeaderFooter = true;
        _opts.EvenPageHeaderFormat = format;
        return this;
    }

    /// <summary>Builds an even-page-only header from text and page tokens.</summary>
    /// <param name="build">Delegate to compose even-page header segments.</param>
    public PdfHeaderCompose EvenPagesText(System.Action<HeaderTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetEvenPageHeaderSegmentsForCompose();
        _opts.ClearEvenPageHeaderZonesForCompose();
        var builder = new HeaderTextBuilder(segments);
        build(builder);
        return this;
    }
}
