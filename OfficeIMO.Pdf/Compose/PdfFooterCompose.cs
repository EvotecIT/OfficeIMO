using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Footer builder (alignment, text, page number tokens).</summary>
public class PdfFooterCompose {
    private readonly PdfOptions _opts;
    internal PdfFooterCompose(PdfOptions opts) { _opts = opts; }
    /// <summary>Sets footer alignment to the left.</summary>
    public PdfFooterCompose AlignLeft() { _opts.FooterAlign = PdfAlign.Left; return this; }
    /// <summary>Sets footer alignment to the center.</summary>
    public PdfFooterCompose AlignCenter() { _opts.FooterAlign = PdfAlign.Center; return this; }
    /// <summary>Sets footer alignment to the right.</summary>
    public PdfFooterCompose AlignRight() { _opts.FooterAlign = PdfAlign.Right; return this; }
    /// <summary>Sets footer text color.</summary>
    public PdfFooterCompose Color(PdfColor color) { _opts.FooterTextColor = color; return this; }
    /// <summary>Sets footer font.</summary>
    public PdfFooterCompose Font(PdfStandardFont font) { _opts.FooterFont = font; _opts.FooterFontFamily = null; return this; }
    /// <summary>Uses a registered named font family for footer text.</summary>
    public PdfFooterCompose FontFamily(string familyName) { _opts.FooterFontFamily = familyName; return this; }
    /// <summary>Sets footer font size in points.</summary>
    public PdfFooterCompose FontSize(double size) { Guard.Positive(size, nameof(size)); _opts.FooterFontSize = size; return this; }
    /// <summary>Sets footer baseline offset below the bottom margin in points.</summary>
    public PdfFooterCompose Offset(double points) { Guard.NonNegative(points, nameof(points)); _opts.FooterOffsetY = points; return this; }
    /// <summary>Renders the current page number in the footer.</summary>
    public PdfFooterCompose PageNumber() { _opts.ClearFooterSegmentsForCompose(); _opts.ClearFooterZonesForCompose(); _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}"; return this; }
    /// <summary>Renders the current page number and total pages in the footer.</summary>
    public PdfFooterCompose PageNumberWithTotal() { _opts.ClearFooterSegmentsForCompose(); _opts.ClearFooterZonesForCompose(); _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}/{pages}"; return this; }
    /// <summary>Renders left, center, and right footer zones on one line. Supports {page} and {pages}.</summary>
    public PdfFooterCompose Zones(string? left, string? center, string? right) {
        _opts.SetFooterZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders page-1-only left, center, and right footer zones. Supports {page} and {pages}.</summary>
    public PdfFooterCompose FirstPageZones(string? left, string? center, string? right) {
        _opts.SetFirstPageFooterZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders even-page-only left, center, and right footer zones. Supports {page} and {pages}.</summary>
    public PdfFooterCompose EvenPagesZones(string? left, string? center, string? right) {
        _opts.SetEvenPageFooterZonesForCompose(left, center, right);
        return this;
    }

    /// <summary>Adds an image to the running footer.</summary>
    public PdfFooterCompose Image(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) =>
        Image(data, width, height, align, fit, alternativeText: null);

    /// <summary>Adds a meaningful image to the running footer with alternate text.</summary>
    public PdfFooterCompose Image(byte[] data, double width, double height, string? alternativeText) =>
        Image(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the running footer.</summary>
    public PdfFooterCompose Image(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        _opts.AddFooterImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds an image to the page-1-only footer.</summary>
    public PdfFooterCompose FirstPageImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) =>
        FirstPageImage(data, width, height, align, fit, alternativeText: null);

    /// <summary>Adds a meaningful image to the page-1-only footer with alternate text.</summary>
    public PdfFooterCompose FirstPageImage(byte[] data, double width, double height, string? alternativeText) =>
        FirstPageImage(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the page-1-only footer.</summary>
    public PdfFooterCompose FirstPageImage(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        _opts.AddFirstPageFooterImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds an image to the even-page-only footer.</summary>
    public PdfFooterCompose EvenPagesImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) =>
        EvenPagesImage(data, width, height, align, fit, alternativeText: null);

    /// <summary>Adds a meaningful image to the even-page-only footer with alternate text.</summary>
    public PdfFooterCompose EvenPagesImage(byte[] data, double width, double height, string? alternativeText) =>
        EvenPagesImage(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the even-page-only footer.</summary>
    public PdfFooterCompose EvenPagesImage(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        _opts.AddEvenPageFooterImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds a shape to the running footer.</summary>
    public PdfFooterCompose Shape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddFooterShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Adds a shape to the page-1-only footer.</summary>
    public PdfFooterCompose FirstPageShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddFirstPageFooterShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Adds a shape to the even-page-only footer.</summary>
    public PdfFooterCompose EvenPagesShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddEvenPageFooterShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Renders a literal footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterCompose Text(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearFooterSegmentsForCompose();
        _opts.ClearFooterZonesForCompose();
        _opts.ShowPageNumbers = true;
        _opts.FooterFormat = format;
        return this;
    }

    /// <summary>Builds a custom footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose footer segments.</param>
    public PdfFooterCompose Text(System.Action<FooterTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetFooterSegmentsForCompose();
        _opts.ClearFooterZonesForCompose();
        var b = new FooterTextBuilder(segments);
        build(b);
        _opts.ShowPageNumbers = true; // might be needed when builder inserts tokens
        return this;
    }

    /// <summary>Renders a page-1-only footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterCompose FirstPageText(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearFirstPageFooterSegmentsForCompose();
        _opts.ClearFirstPageFooterZonesForCompose();
        _opts.DifferentFirstPageHeaderFooter = true;
        _opts.FirstPageFooterFormat = format;
        return this;
    }

    /// <summary>Builds a page-1-only footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose first-page footer segments.</param>
    public PdfFooterCompose FirstPageText(System.Action<FooterTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetFirstPageFooterSegmentsForCompose();
        _opts.ClearFirstPageFooterZonesForCompose();
        var b = new FooterTextBuilder(segments);
        build(b);
        return this;
    }

    /// <summary>Renders an even-page-only footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterCompose EvenPagesText(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearEvenPageFooterSegmentsForCompose();
        _opts.ClearEvenPageFooterZonesForCompose();
        _opts.DifferentOddAndEvenPagesHeaderFooter = true;
        _opts.EvenPageFooterFormat = format;
        return this;
    }

    /// <summary>Builds an even-page-only footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose even-page footer segments.</param>
    public PdfFooterCompose EvenPagesText(System.Action<FooterTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetEvenPageFooterSegmentsForCompose();
        _opts.ClearEvenPageFooterZonesForCompose();
        var b = new FooterTextBuilder(segments);
        build(b);
        return this;
    }
}
