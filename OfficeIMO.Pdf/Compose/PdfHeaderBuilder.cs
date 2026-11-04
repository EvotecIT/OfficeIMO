using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Header builder. Forwards to the header slots of <see cref="PdfPageSetup"/>.</summary>
public sealed partial class PdfHeaderBuilder {
    private readonly PdfCreateOptions _opts;
    private readonly PdfDocument _doc;
    internal PdfHeaderBuilder(PdfDocument doc, PdfCreateOptions opts) { _doc = doc; _opts = opts; }
    private PdfHeaderFooter Running => _opts.PageSetup.Header;
    private PdfHeaderFooter FirstPage => _opts.PageSetup.FirstPageHeader;
    private PdfHeaderFooter EvenPages => _opts.PageSetup.EvenPageHeader;

    /// <summary>Sets <see cref="PdfHeaderFooter.Alignment"/> of the header text.</summary>
    public PdfHeaderBuilder Alignment(OfficeAlignment alignment) { Running.Alignment = alignment; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.Color"/> of the header text.</summary>
    public PdfHeaderBuilder Color(global::OfficeIMO.OfficeColor color) { Running.Color = color; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.StandardFont"/> and clears <see cref="PdfHeaderFooter.FontFamily"/>.</summary>
    public PdfHeaderBuilder StandardFont(PdfStandardFont font) { Running.StandardFont = font; Running.FontFamily = null; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.FontFamily"/> to a registered named font family.</summary>
    public PdfHeaderBuilder FontFamily(string familyName) { Running.FontFamily = familyName; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.FontSize"/>.</summary>
    public PdfHeaderBuilder FontSize(global::OfficeIMO.OfficeLength size) { Guard.Positive(size.ToPoints(), nameof(size)); Running.FontSize = size; return this; }
    /// <summary>Sets <see cref="PdfHeaderFooter.Offset"/>, the baseline distance above the top margin.</summary>
    public PdfHeaderBuilder Offset(global::OfficeIMO.OfficeLength offset) { Guard.NonNegative(offset.ToPoints(), nameof(offset)); Running.Offset = offset; return this; }
    /// <summary>Renders the current page number in the header.</summary>
    public PdfHeaderBuilder PageNumber() => Text("{page}");
    /// <summary>Renders the current page number and total pages in the header.</summary>
    public PdfHeaderBuilder PageNumberWithTotal() => Text("{page}/{pages}");
    /// <summary>Renders left, center, and right header zones on one line. Supports {page} and {pages}.</summary>
    public PdfHeaderBuilder Zones(string? left, string? center, string? right) { Running.SetZones(left, center, right); return this; }
    /// <summary>Renders independently styled left, center, and right header zones on one line.</summary>
    public PdfHeaderBuilder StyledZones(System.Action<PdfHeaderTextBuilder>? left, System.Action<PdfHeaderTextBuilder>? center, System.Action<PdfHeaderTextBuilder>? right) {
        Running.SetStyledZones(BuildSegments(left), BuildSegments(center), BuildSegments(right));
        return this;
    }
    /// <summary>Renders page-1-only left, center, and right header zones. Supports {page} and {pages}.</summary>
    public PdfHeaderBuilder FirstPageZones(string? left, string? center, string? right) { FirstPage.SetZones(left, center, right); return this; }
    /// <summary>Renders independently styled page-1-only header zones.</summary>
    public PdfHeaderBuilder FirstPageStyledZones(System.Action<PdfHeaderTextBuilder>? left, System.Action<PdfHeaderTextBuilder>? center, System.Action<PdfHeaderTextBuilder>? right) {
        FirstPage.SetStyledZones(BuildSegments(left), BuildSegments(center), BuildSegments(right));
        return this;
    }
    /// <summary>Renders even-page-only left, center, and right header zones. Supports {page} and {pages}.</summary>
    public PdfHeaderBuilder EvenPagesZones(string? left, string? center, string? right) { EvenPages.SetZones(left, center, right); return this; }
    /// <summary>Renders independently styled even-page-only header zones.</summary>
    public PdfHeaderBuilder EvenPagesStyledZones(System.Action<PdfHeaderTextBuilder>? left, System.Action<PdfHeaderTextBuilder>? center, System.Action<PdfHeaderTextBuilder>? right) {
        EvenPages.SetStyledZones(BuildSegments(left), BuildSegments(center), BuildSegments(right));
        return this;
    }

    private static System.Collections.Generic.List<PdfFooterSegment>? BuildSegments(System.Action<PdfHeaderTextBuilder>? build) {
        if (build == null) return null;
        var segments = new System.Collections.Generic.List<PdfFooterSegment>();
        build(new PdfHeaderTextBuilder(segments));
        return segments;
    }

    /// <summary>Adds an image to the running header.</summary>
    public PdfHeaderBuilder Image(PdfHeaderFooterImage image) { Running.AddImage(image); return this; }
    /// <summary>Adds an image with its reusable style to the page-1-only header.</summary>
    public PdfHeaderBuilder FirstPageImage(PdfHeaderFooterImage image) { FirstPage.AddImage(image); return this; }
    /// <summary>Adds an image with its reusable style to the even-page-only header.</summary>
    public PdfHeaderBuilder EvenPagesImage(PdfHeaderFooterImage image) { EvenPages.AddImage(image); return this; }

    /// <summary>Adds a meaningful image to the running header with alternate text.</summary>
    public PdfHeaderBuilder Image(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        Image(data, width, height, OfficeAlignment.Left, OfficeImageFit.Stretch, alternativeText);
    /// <summary>Adds an image to the running header.</summary>
    public PdfHeaderBuilder Image(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment align = OfficeAlignment.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) =>
        Image(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
    /// <summary>Adds a meaningful image to the page-1-only header with alternate text.</summary>
    public PdfHeaderBuilder FirstPageImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        FirstPageImage(data, width, height, OfficeAlignment.Left, OfficeImageFit.Stretch, alternativeText);
    /// <summary>Adds an image to the page-1-only header.</summary>
    public PdfHeaderBuilder FirstPageImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment align = OfficeAlignment.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) =>
        FirstPageImage(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));
    /// <summary>Adds a meaningful image to the even-page-only header with alternate text.</summary>
    public PdfHeaderBuilder EvenPagesImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText) =>
        EvenPagesImage(data, width, height, OfficeAlignment.Left, OfficeImageFit.Stretch, alternativeText);
    /// <summary>Adds an image to the even-page-only header.</summary>
    public PdfHeaderBuilder EvenPagesImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment align = OfficeAlignment.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null) =>
        EvenPagesImage(new PdfHeaderFooterImage(data, width, height, align, fit, alternativeText));

    /// <summary>Adds a shape to the running header.</summary>
    public PdfHeaderBuilder Shape(OfficeShape shape, OfficeAlignment align = OfficeAlignment.Left) { Running.AddShape(new PdfHeaderFooterShape(shape, align)); return this; }
    /// <summary>Adds a shape to the page-1-only header.</summary>
    public PdfHeaderBuilder FirstPageShape(OfficeShape shape, OfficeAlignment align = OfficeAlignment.Left) { FirstPage.AddShape(new PdfHeaderFooterShape(shape, align)); return this; }
    /// <summary>Adds a shape to the even-page-only header.</summary>
    public PdfHeaderBuilder EvenPagesShape(OfficeShape shape, OfficeAlignment align = OfficeAlignment.Left) { EvenPages.AddShape(new PdfHeaderFooterShape(shape, align)); return this; }

    /// <summary>Renders a literal header text format. Supports {page} and {pages}.</summary>
    public PdfHeaderBuilder Text(string format) { Running.SetText(format); return this; }
    /// <summary>Builds a custom header from text and page tokens.</summary>
    /// <param name="build">Delegate to compose header segments.</param>
    public PdfHeaderBuilder Text(System.Action<PdfHeaderTextBuilder> build) { Guard.NotNull(build, nameof(build)); Running.SetText(BuildSegments(build)!); return this; }
    /// <summary>Renders a page-1-only header text format. Supports {page} and {pages}.</summary>
    public PdfHeaderBuilder FirstPageText(string format) { FirstPage.SetText(format); return this; }
    /// <summary>Builds a page-1-only header from text and page tokens.</summary>
    /// <param name="build">Delegate to compose first-page header segments.</param>
    public PdfHeaderBuilder FirstPageText(System.Action<PdfHeaderTextBuilder> build) { Guard.NotNull(build, nameof(build)); FirstPage.SetText(BuildSegments(build)!); return this; }
    /// <summary>Renders an even-page-only header text format. Supports {page} and {pages}.</summary>
    public PdfHeaderBuilder EvenPagesText(string format) { EvenPages.SetText(format); return this; }
    /// <summary>Builds an even-page-only header from text and page tokens.</summary>
    /// <param name="build">Delegate to compose even-page header segments.</param>
    public PdfHeaderBuilder EvenPagesText(System.Action<PdfHeaderTextBuilder> build) { Guard.NotNull(build, nameof(build)); EvenPages.SetText(BuildSegments(build)!); return this; }
}
