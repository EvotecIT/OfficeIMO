namespace OfficeIMO.Pdf;

/// <summary>Rectangle requested for redaction planning, using PDF point coordinates from the page bottom-left.</summary>
/// <remarks>Native page coordinates use default user-space units: one unit is /UserUnit physical points (1/72 inch each), defaulting to 1. Local text/font and matrix values retain their native coordinate spaces. Explicit visual projections use their documented coordinates.</remarks>
public sealed class PdfRedactionArea {
    /// <summary>Creates a redaction area.</summary>
    public PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label = null)
        : this(pageNumber, x, y, width, height, label, textRenderingMode: null, exactGeometry: null, PdfRedactionContentScope.TextAndUnderlay, PdfRedactionAppearanceMode.Exact) {
    }

    /// <summary>Creates a redaction area with an explicit intersecting-content policy.</summary>
    public PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label, PdfRedactionContentScope contentScope)
        : this(pageNumber, x, y, width, height, label, textRenderingMode: null, exactGeometry: null, contentScope, PdfRedactionAppearanceMode.Exact) {
    }

    /// <summary>Creates a redaction area with an explicit visible-appearance policy.</summary>
    public PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label, PdfRedactionAppearanceMode appearanceMode)
        : this(pageNumber, x, y, width, height, label, textRenderingMode: null, exactGeometry: null, PdfRedactionContentScope.TextAndUnderlay, appearanceMode) {
    }

    /// <summary>Creates a redaction area with explicit intersecting-content and visible-appearance policies.</summary>
    public PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label, PdfRedactionContentScope contentScope, PdfRedactionAppearanceMode appearanceMode)
        : this(pageNumber, x, y, width, height, label, textRenderingMode: null, exactGeometry: null, contentScope, appearanceMode) {
    }

    private PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label, int? textRenderingMode, PdfRedactionGeometry? exactGeometry, PdfRedactionContentScope contentScope, PdfRedactionAppearanceMode appearanceMode, PdfRedactionGlyphSelection? glyphSelection = null, PdfRedactionSearchCriteria? searchCriteria = null) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number must be greater than zero.");
        }

        if (!IsFinite(x)) {
            throw new ArgumentOutOfRangeException(nameof(x), "X coordinate must be finite.");
        }

        if (!IsFinite(y)) {
            throw new ArgumentOutOfRangeException(nameof(y), "Y coordinate must be finite.");
        }

        if (!IsFinite(width) || width <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be finite and greater than zero.");
        }

        if (!IsFinite(height) || height <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(height), "Height must be finite and greater than zero.");
        }
        if (contentScope is < PdfRedactionContentScope.TextOnly or > PdfRedactionContentScope.TextAndUnderlay) throw new ArgumentOutOfRangeException(nameof(contentScope));
        if (appearanceMode is < PdfRedactionAppearanceMode.Exact or > PdfRedactionAppearanceMode.FullLine) throw new ArgumentOutOfRangeException(nameof(appearanceMode));

        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Label = label;
        ContentScope = contentScope;
        AppearanceMode = appearanceMode;
        TextRenderingMode = textRenderingMode;
        ExactGeometry = exactGeometry;
        GlyphSelection = glyphSelection;
        SearchCriteria = searchCriteria;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Left coordinate in page user-space units.</summary>
    public double X { get; }

    /// <summary>Bottom coordinate in page user-space units.</summary>
    public double Y { get; }

    /// <summary>Rectangle width in page user-space units.</summary>
    public double Width { get; }

    /// <summary>Rectangle height in page user-space units.</summary>
    public double Height { get; }

    /// <summary>Optional caller label.</summary>
    public string? Label { get; }

    /// <summary>Intersecting content removal policy for this reviewed area.</summary>
    public PdfRedactionContentScope ContentScope { get; }

    /// <summary>Visible privacy-appearance policy for this reviewed area.</summary>
    public PdfRedactionAppearanceMode AppearanceMode { get; }

    /// <summary>Right coordinate in page user-space units.</summary>
    public double Right => X + Width;

    /// <summary>Top coordinate in page user-space units.</summary>
    public double Top => Y + Height;

    /// <summary>Returns this area with a different review label, preserving its geometry and removal policy.</summary>
    /// <remarks>Use this when editing a search-derived mark. Reconstructing a rectangle loses its precise glyph and rotated geometry evidence.</remarks>
    public PdfRedactionArea WithLabel(string? label) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, label, TextRenderingMode, ExactGeometry, ContentScope, AppearanceMode, GlyphSelection, SearchCriteria);

    internal int? TextRenderingMode { get; }

    internal PdfRedactionGeometry? ExactGeometry { get; }

    // A search mark removes only the glyphs it selected. The rectangle stays the visible mark and the
    // underlay boundary, and it must never silently expand removal to an entire text object.
    internal PdfRedactionGlyphSelection? GlyphSelection { get; }

    internal bool RequiresGlyphRewrite => GlyphSelection is not null;

    internal bool IsHiddenOnly => GlyphSelection?.IsHiddenOnly == true;

    // What the search looked for, kept so copies outside the page content can be cleared and checked.
    internal PdfRedactionSearchCriteria? SearchCriteria { get; }

    /// <summary>Decides whether this area removes one glyph: by identity for search marks, by geometry otherwise.</summary>
    internal bool TargetsGlyph(PdfTextSpanBounds bounds, double originX, double originY, IReadOnlyList<byte>? encodedBytes) =>
        IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height) &&
        (GlyphSelection is null || encodedBytes is not null && GlyphSelection.Contains(originX, originY, encodedBytes));

    /// <summary>Decides whether a text run can hold a glyph this area removes.</summary>
    internal bool TargetsSpan(PdfTextSpanBounds bounds) =>
        IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height) &&
        (GlyphSelection is null || GlyphSelection.HasOriginWithin(bounds));

    internal bool IntersectsRectangle(double x, double y, double width, double height) =>
        ExactGeometry?.IntersectsRectangle(x, y, width, height) ??
        (RequiresGlyphRewrite
            // Independently accumulated advances can differ at a shared edge by floating-point noise.
            ? X < x + width - 0.0000001D && Right > x + 0.0000001D && Y < y + height - 0.0000001D && Top > y + 0.0000001D
            : X < x + width && Right > x && Y < y + height && Top > y);

    internal bool ContainsRectangle(double x, double y, double width, double height) =>
        ExactGeometry?.ContainsRectangle(x, y, width, height) ??
        x >= X && x + width <= Right && y >= Y && y + height <= Top;

    internal bool ContainsPoint(double x, double y) =>
        ExactGeometry?.ContainsPoint(x, y) ?? x >= X && x <= Right && y >= Y && y <= Top;

    internal bool IntersectsQuadrilateral(
        PdfRedactionPoint first,
        PdfRedactionPoint second,
        PdfRedactionPoint third,
        PdfRedactionPoint fourth) =>
        ExactGeometry?.IntersectsQuadrilateral(first, second, third, fourth) ??
        PdfRedactionGeometry.RectangleIntersectsQuadrilateral(X, Y, Width, Height, first, second, third, fourth);

    internal PdfRedactionArea WithExactGeometry(PdfRedactionGeometry exactGeometry) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, TextRenderingMode, exactGeometry, ContentScope, AppearanceMode, GlyphSelection, SearchCriteria);

    internal PdfRedactionArea WithTextRenderingMode(int textRenderingMode) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, textRenderingMode, ExactGeometry, ContentScope, AppearanceMode, GlyphSelection, SearchCriteria);

    internal PdfRedactionArea WithPolicies(PdfRedactionContentScope contentScope, PdfRedactionAppearanceMode appearanceMode) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, TextRenderingMode, ExactGeometry, contentScope, appearanceMode, GlyphSelection, SearchCriteria);

    internal PdfRedactionArea WithSearchCriteria(PdfRedactionSearchCriteria? searchCriteria) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, TextRenderingMode, ExactGeometry, ContentScope, AppearanceMode, GlyphSelection, searchCriteria);

    internal PdfRedactionArea WithGlyphSelection(PdfRedactionGlyphSelection glyphSelection) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, TextRenderingMode, ExactGeometry, ContentScope, AppearanceMode, glyphSelection, SearchCriteria);

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
