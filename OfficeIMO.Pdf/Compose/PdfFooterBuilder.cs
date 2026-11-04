using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Footer builder. Forwards to the footer slots of <see cref="PdfPageSetup"/>.</summary>
public sealed partial class PdfFooterBuilder {
    private readonly PdfCreateOptions _opts;
    private readonly PdfDocument _doc;
    internal PdfFooterBuilder(PdfDocument doc, PdfCreateOptions opts) { _doc = doc; _opts = opts; }
    private PdfHeaderFooter Running => _opts.PageSetup.Footer;
    private PdfHeaderFooter FirstPage => _opts.PageSetup.FirstPageFooter;
    private PdfHeaderFooter EvenPages => _opts.PageSetup.EvenPageFooter;

    /// <summary>Sets <see cref="PdfHeaderFooter.Alignment"/> of the footer text.</summary>
    public PdfFooterBuilder Alignment(OfficeAlignment alignment) { Running.Alignment = alignment; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.Color"/> of the footer text.</summary>
    public PdfFooterBuilder Color(global::OfficeIMO.OfficeColor color) { Running.Color = color; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.StandardFont"/> and clears <see cref="PdfHeaderFooter.FontFamily"/>.</summary>
    public PdfFooterBuilder StandardFont(PdfStandardFont font) { Running.StandardFont = font; Running.FontFamily = null; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.FontFamily"/> to a registered named font family.</summary>
    public PdfFooterBuilder FontFamily(string familyName) { Running.FontFamily = familyName; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.FontSize"/>.</summary>
    public PdfFooterBuilder FontSize(global::OfficeIMO.OfficeLength size) { Guard.Positive(size.ToPoints(), nameof(size)); Running.FontSize = size; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.Offset"/>, the baseline distance below the bottom margin.</summary>
    public PdfFooterBuilder Offset(global::OfficeIMO.OfficeLength offset) { Guard.NonNegative(offset.ToPoints(), nameof(offset)); Running.Offset = offset; return this; }
    /// <summary>Renders the current page number in the footer.</summary>
    public PdfFooterBuilder PageNumber() => Text("{page}");
    /// <summary>Renders the current page number and total pages in the footer.</summary>
    public PdfFooterBuilder PageNumberWithTotal() => Text("{page}/{pages}");
    /// <summary>Renders left, center, and right footer zones on one line. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder Zones(string? left, string? center, string? right) { Running.SetZones(left, center, right); return this; }
    /// <summary>Renders independently styled left, center, and right footer zones on one line.</summary>
    public PdfFooterBuilder StyledZones(System.Action<PdfFooterTextBuilder>? left, System.Action<PdfFooterTextBuilder>? center, System.Action<PdfFooterTextBuilder>? right) {
        Running.SetStyledZones(BuildSegments(left), BuildSegments(center), BuildSegments(right));
        return this;
    }
    /// <summary>Renders page-1-only left, center, and right footer zones. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder FirstPageZones(string? left, string? center, string? right) { FirstPage.SetZones(left, center, right); return this; }
    /// <summary>Renders independently styled page-1-only footer zones.</summary>
    public PdfFooterBuilder FirstPageStyledZones(System.Action<PdfFooterTextBuilder>? left, System.Action<PdfFooterTextBuilder>? center, System.Action<PdfFooterTextBuilder>? right) {
        FirstPage.SetStyledZones(BuildSegments(left), BuildSegments(center), BuildSegments(right));
        return this;
    }
    /// <summary>Renders even-page-only left, center, and right footer zones. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder EvenPagesZones(string? left, string? center, string? right) { EvenPages.SetZones(left, center, right); return this; }
    /// <summary>Renders independently styled even-page-only footer zones.</summary>
    public PdfFooterBuilder EvenPagesStyledZones(System.Action<PdfFooterTextBuilder>? left, System.Action<PdfFooterTextBuilder>? center, System.Action<PdfFooterTextBuilder>? right) {
        EvenPages.SetStyledZones(BuildSegments(left), BuildSegments(center), BuildSegments(right));
        return this;
    }

    private static System.Collections.Generic.List<PdfFooterSegment>? BuildSegments(System.Action<PdfFooterTextBuilder>? build) {
        if (build == null) return null;
        var segments = new System.Collections.Generic.List<PdfFooterSegment>();
        build(new PdfFooterTextBuilder(segments));
        return segments;
    }

    /// <summary>Adds an image to the running footer.</summary>
    public PdfFooterBuilder Image(PdfHeaderFooterImage image) { Running.AddImage(image); return this; }
    /// <summary>Adds an image with its reusable style to the page-1-only footer.</summary>
    public PdfFooterBuilder FirstPageImage(PdfHeaderFooterImage image) { FirstPage.AddImage(image); return this; }
    /// <summary>Adds an image with its reusable style to the even-page-only footer.</summary>
    public PdfFooterBuilder EvenPagesImage(PdfHeaderFooterImage image) { EvenPages.AddImage(image); return this; }

    /// <summary>Adds a meaningful image to the running footer with alternate text.</summary>
    public PdfFooterBuilder Image(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        Image(data, width, height, OfficeAlignment.Left, OfficeImageFit.Stretch, alternativeText);
    /// <summary>Adds an image to the running footer.</summary>
    public PdfFooterBuilder Image(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment align = OfficeAlignment.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) =>
        Image(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
    /// <summary>Adds a meaningful image to the page-1-only footer with alternate text.</summary>
    public PdfFooterBuilder FirstPageImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        FirstPageImage(data, width, height, OfficeAlignment.Left, OfficeImageFit.Stretch, alternativeText);
    /// <summary>Adds an image to the page-1-only footer.</summary>
    public PdfFooterBuilder FirstPageImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment align = OfficeAlignment.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) =>
        FirstPageImage(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
    /// <summary>Adds a meaningful image to the even-page-only footer with alternate text.</summary>
    public PdfFooterBuilder EvenPagesImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        EvenPagesImage(data, width, height, OfficeAlignment.Left, OfficeImageFit.Stretch, alternativeText);
    /// <summary>Adds an image to the even-page-only footer.</summary>
    public PdfFooterBuilder EvenPagesImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment align = OfficeAlignment.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) =>
        EvenPagesImage(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));

    /// <summary>Adds a shape to the running footer.</summary>
    public PdfFooterBuilder Shape(OfficeShape shape, OfficeAlignment align = OfficeAlignment.Left) { Running.AddShape(new PdfHeaderFooterShape(shape, align)); return this; }
    /// <summary>Adds a shape to the page-1-only footer.</summary>
    public PdfFooterBuilder FirstPageShape(OfficeShape shape, OfficeAlignment align = OfficeAlignment.Left) { FirstPage.AddShape(new PdfHeaderFooterShape(shape, align)); return this; }
    /// <summary>Adds a shape to the even-page-only footer.</summary>
    public PdfFooterBuilder EvenPagesShape(OfficeShape shape, OfficeAlignment align = OfficeAlignment.Left) { EvenPages.AddShape(new PdfHeaderFooterShape(shape, align)); return this; }

    /// <summary>Renders a literal footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder Text(string format) { Running.SetText(format); return this; }
    /// <summary>Builds a custom footer from text and page tokens.</summary>
    /// <param name="build">Delegate to compose footer segments.</param>
    public PdfFooterBuilder Text(System.Action<PdfFooterTextBuilder> build) { Guard.NotNull(build, nameof(build)); Running.SetText(BuildSegments(build)!); return this; }
    /// <summary>Renders a page-1-only footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder FirstPageText(string format) { FirstPage.SetText(format); return this; }
    /// <summary>Builds a page-1-only footer from text and page tokens.</summary>
    /// <param name="build">Delegate to compose first-page footer segments.</param>
    public PdfFooterBuilder FirstPageText(System.Action<PdfFooterTextBuilder> build) { Guard.NotNull(build, nameof(build)); FirstPage.SetText(BuildSegments(build)!); return this; }
    /// <summary>Renders an even-page-only footer text format. Supports {page} and {pages}.</summary>
    public PdfFooterBuilder EvenPagesText(string format) { EvenPages.SetText(format); return this; }
    /// <summary>Builds an even-page-only footer from text and page tokens.</summary>
    /// <param name="build">Delegate to compose even-page footer segments.</param>
    public PdfFooterBuilder EvenPagesText(System.Action<PdfFooterTextBuilder> build) { Guard.NotNull(build, nameof(build)); EvenPages.SetText(BuildSegments(build)!); return this; }
}
