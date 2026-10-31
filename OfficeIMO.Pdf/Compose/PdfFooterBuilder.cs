using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Footer builder (alignment, text, page number tokens).</summary>
public sealed partial class PdfFooterBuilder {
    private readonly PdfCreateOptions _opts;
    private readonly PdfDocument _doc;
    internal PdfFooterBuilder(PdfDocument doc, PdfCreateOptions opts) { _doc = doc; _opts = opts; }
    /// <summary>Sets footer alignment to the left.</summary>
    public PdfFooterBuilder AlignLeft() { _opts.FooterAlign = PdfAlign.Left; return this; }
    /// <summary>Sets footer alignment to the center.</summary>
    public PdfFooterBuilder AlignCenter() { _opts.FooterAlign = PdfAlign.Center; return this; }
    /// <summary>Sets footer alignment to the right.</summary>
    public PdfFooterBuilder AlignRight() { _opts.FooterAlign = PdfAlign.Right; return this; }
    /// <summary>Sets footer text color.</summary>
    public PdfFooterBuilder Color(global::OfficeIMO.OfficeColor color) { _opts.FooterTextColor = color; return this; }
    /// <summary>Sets footer font.</summary>
    public PdfFooterBuilder Font(PdfStandardFont font) { _opts.FooterFont = font; _opts.FooterFontFamily = null; return this; }
    /// <summary>Uses a registered named font family for footer text.</summary>
    public PdfFooterBuilder FontFamily(string familyName) { _opts.FooterFontFamily = familyName; return this; }
    /// <summary>Sets footer font size in points.</summary>
    public PdfFooterBuilder FontSize(global::OfficeIMO.OfficeLength size) { Guard.Positive(size.ToPoints(), nameof(size)); _opts.FooterFontSize = size; return this; }
    /// <summary>Sets footer baseline offset below the bottom margin in points.</summary>
    public PdfFooterBuilder Offset(global::OfficeIMO.OfficeLength points) { Guard.NonNegative(points.ToPoints(), nameof(points)); _opts.FooterOffsetY = points; return this; }
    /// <summary>Renders the current page number in the footer.</summary>
    public PdfFooterBuilder PageNumber() { _opts.ClearFooterSegmentsForCompose(); _opts.ClearFooterZonesForCompose(); _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}"; return this; }
    /// <summary>Renders the current page number and total pages in the footer.</summary>
    public PdfFooterBuilder PageNumberWithTotal() { _opts.ClearFooterSegmentsForCompose(); _opts.ClearFooterZonesForCompose(); _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}/{pages}"; return this; }
    /// <summary>Renders left, center, and right footer zones on one line. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder Zones(string? left, string? center, string? right) {
        _opts.SetFooterZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders independently styled left, center, and right footer zones on one line.</summary>
    public PdfFooterBuilder StyledZones(System.Action<PdfFooterTextBuilder>? left, System.Action<PdfFooterTextBuilder>? center, System.Action<PdfFooterTextBuilder>? right) {
        _opts.SetFooterZoneSegmentsForCompose(BuildZone(left), BuildZone(center), BuildZone(right));
        return this;
    }
    /// <summary>Renders page-1-only left, center, and right footer zones. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder FirstPageZones(string? left, string? center, string? right) {
        _opts.SetFirstPageFooterZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders independently styled page-1-only footer zones.</summary>
    public PdfFooterBuilder FirstPageStyledZones(System.Action<PdfFooterTextBuilder>? left, System.Action<PdfFooterTextBuilder>? center, System.Action<PdfFooterTextBuilder>? right) {
        _opts.SetFirstPageFooterZoneSegmentsForCompose(BuildZone(left), BuildZone(center), BuildZone(right));
        return this;
    }
    /// <summary>Renders even-page-only left, center, and right footer zones. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder EvenPagesZones(string? left, string? center, string? right) {
        _opts.SetEvenPageFooterZonesForCompose(left, center, right);
        return this;
    }
    /// <summary>Renders independently styled even-page-only footer zones.</summary>
    public PdfFooterBuilder EvenPagesStyledZones(System.Action<PdfFooterTextBuilder>? left, System.Action<PdfFooterTextBuilder>? center, System.Action<PdfFooterTextBuilder>? right) {
        _opts.SetEvenPageFooterZoneSegmentsForCompose(BuildZone(left), BuildZone(center), BuildZone(right));
        return this;
    }

    private static System.Collections.Generic.List<PdfFooterSegment>? BuildZone(System.Action<PdfFooterTextBuilder>? build) {
        if (build == null) return null;
        var segments = new System.Collections.Generic.List<PdfFooterSegment>();
        build(new PdfFooterTextBuilder(segments));
        return segments;
    }

    /// <summary>Adds an image to the running footer.</summary>
    public PdfFooterBuilder Image(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        _opts.AddFooterImageForCompose(image);
        return this;
    }

    /// <summary>Adds an image with its reusable style to the page-1-only footer.</summary>
    public PdfFooterBuilder FirstPageImage(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        _opts.AddFirstPageFooterImageForCompose(image);
        return this;
    }

    /// <summary>Adds an image with its reusable style to the even-page-only footer.</summary>
    public PdfFooterBuilder EvenPagesImage(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        _opts.AddEvenPageFooterImageForCompose(image);
        return this;
    }

    /// <summary>Adds a meaningful image to the running footer with alternate text.</summary>
    public PdfFooterBuilder Image(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        Image(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the running footer.</summary>
    public PdfFooterBuilder Image(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) {
        _opts.AddFooterImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds a meaningful image to the page-1-only footer with alternate text.</summary>
    public PdfFooterBuilder FirstPageImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        FirstPageImage(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the page-1-only footer.</summary>
    public PdfFooterBuilder FirstPageImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) {
        _opts.AddFirstPageFooterImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds a meaningful image to the even-page-only footer with alternate text.</summary>
    public PdfFooterBuilder EvenPagesImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        EvenPagesImage(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText);

    /// <summary>Adds an image to the even-page-only footer.</summary>
    public PdfFooterBuilder EvenPagesImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) {
        _opts.AddEvenPageFooterImageForCompose(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
        return this;
    }

    /// <summary>Adds a shape to the running footer.</summary>
    public PdfFooterBuilder Shape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddFooterShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Adds a shape to the page-1-only footer.</summary>
    public PdfFooterBuilder FirstPageShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddFirstPageFooterShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Adds a shape to the even-page-only footer.</summary>
    public PdfFooterBuilder EvenPagesShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        _opts.AddEvenPageFooterShapeForCompose(new PdfHeaderFooterShape(shape, align));
        return this;
    }

    /// <summary>Renders a literal footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder Text(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearFooterSegmentsForCompose();
        _opts.ClearFooterZonesForCompose();
        _opts.ShowPageNumbers = true;
        _opts.FooterFormat = format;
        return this;
    }

    /// <summary>Builds a custom footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose footer segments.</param>
    public PdfFooterBuilder Text(System.Action<PdfFooterTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetFooterSegmentsForCompose();
        _opts.ClearFooterZonesForCompose();
        var b = new PdfFooterTextBuilder(segments);
        build(b);
        _opts.ShowPageNumbers = true; // might be needed when builder inserts tokens
        return this;
    }

    /// <summary>Renders a page-1-only footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder FirstPageText(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearFirstPageFooterSegmentsForCompose();
        _opts.ClearFirstPageFooterZonesForCompose();
        _opts.DifferentFirstPageHeaderFooter = true;
        _opts.FirstPageFooterFormat = format;
        return this;
    }

    /// <summary>Builds a page-1-only footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose first-page footer segments.</param>
    public PdfFooterBuilder FirstPageText(System.Action<PdfFooterTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetFirstPageFooterSegmentsForCompose();
        _opts.ClearFirstPageFooterZonesForCompose();
        var b = new PdfFooterTextBuilder(segments);
        build(b);
        return this;
    }

    /// <summary>Renders an even-page-only footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder EvenPagesText(string format) {
        Guard.NotNull(format, nameof(format));
        _opts.ClearEvenPageFooterSegmentsForCompose();
        _opts.ClearEvenPageFooterZonesForCompose();
        _opts.DifferentOddAndEvenPagesHeaderFooter = true;
        _opts.EvenPageFooterFormat = format;
        return this;
    }

    /// <summary>Builds an even-page-only footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose even-page footer segments.</param>
    public PdfFooterBuilder EvenPagesText(System.Action<PdfFooterTextBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var segments = _opts.ResetEvenPageFooterSegmentsForCompose();
        _opts.ClearEvenPageFooterZonesForCompose();
        var b = new PdfFooterTextBuilder(segments);
        build(b);
        return this;
    }
}
