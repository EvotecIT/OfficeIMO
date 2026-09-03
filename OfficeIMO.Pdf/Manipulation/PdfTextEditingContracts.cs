using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

/// <summary>A rectangular page region in PDF user space, measured from the page bottom-left.</summary>
public sealed class PdfPageRegion {
    /// <summary>Creates a page region.</summary>
    public PdfPageRegion(int pageNumber, double x, double y, double width, double height) {
        if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number must be greater than zero.");
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Left coordinate in PDF points.</summary>
    public double X { get; }
    /// <summary>Bottom coordinate in PDF points.</summary>
    public double Y { get; }
    /// <summary>Width in PDF points.</summary>
    public double Width { get; }
    /// <summary>Height in PDF points.</summary>
    public double Height { get; }
    /// <summary>Right coordinate in PDF points.</summary>
    public double Right => X + Width;
    /// <summary>Top coordinate in PDF points.</summary>
    public double Top => Y + Height;

    internal PdfRedactionArea ToRedactionArea() => new PdfRedactionArea(PageNumber, X, Y, Width, Height);

    private static void ValidateFinite(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name, "Coordinate must be finite.");
    }

    private static void ValidatePositive(double value, string name) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name, "Dimension must be finite and greater than zero.");
    }
}

/// <summary>Style overrides for existing-page text editing. Null values preserve detected style when possible.</summary>
public sealed class PdfTextEditOptions {
    private PdfStandardFont? _font;
    private double? _fontSize;
    private double? _rotationDegrees;

    /// <summary>Standard PDF font. Null preserves the detected family/style or uses Helvetica for new text.</summary>
    public PdfStandardFont? Font {
        get => _font;
        set {
            if (value.HasValue) Guard.StandardFont(value.Value, nameof(Font), "Text edit font must be a supported standard PDF font.");
            _font = value;
        }
    }

    /// <summary>Font size in points. Null preserves the detected size or uses 12 points for new text.</summary>
    public double? FontSize {
        get => _fontSize;
        set {
            if (value.HasValue && (value.Value <= 0D || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) throw new ArgumentOutOfRangeException(nameof(FontSize), "Text edit font size must be finite and greater than zero.");
            _fontSize = value;
        }
    }

    /// <summary>Text color. Null preserves the detected color or uses black for new text.</summary>
    public PdfColor? Color { get; set; }

    /// <summary>Baseline rotation in degrees. Null preserves the detected rotation or uses zero for new text.</summary>
    public double? RotationDegrees {
        get => _rotationDegrees;
        set {
            if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value))) throw new ArgumentOutOfRangeException(nameof(RotationDegrees), "Text edit rotation must be finite.");
            _rotationDegrees = value;
        }
    }

    /// <summary>
    /// Allows editing non-Type3 text painted with PDF rendering mode 3 while preserving that
    /// invisible rendering mode. Other invisible or clipping text modes remain unsupported.
    /// </summary>
    public bool AllowTextRenderingMode3 { get; set; }

    internal PdfTextEditOptions Snapshot() => new PdfTextEditOptions {
        Font = Font,
        FontSize = FontSize,
        Color = Color,
        RotationDegrees = RotationDegrees,
        AllowTextRenderingMode3 = AllowTextRenderingMode3
    };
}

/// <summary>Text and dominant style detected inside a page region.</summary>
public sealed class PdfRegionText {
    internal PdfRegionText(string text, IReadOnlyList<PdfTextSpan> spans, PdfStandardFont suggestedFont, string? sourceFont, double fontSize, PdfColor color, double rotationDegrees, double baselineX, double baselineY) {
        Text = text;
        Spans = new ReadOnlyCollection<PdfTextSpan>(spans.ToArray());
        SuggestedFont = suggestedFont;
        SourceFont = sourceFont;
        FontSize = fontSize;
        Color = color;
        RotationDegrees = rotationDegrees;
        BaselineX = baselineX;
        BaselineY = baselineY;
    }

    /// <summary>Text assembled in visual line order.</summary>
    public string Text { get; }
    /// <summary>Source spans selected by the region.</summary>
    public IReadOnlyList<PdfTextSpan> Spans { get; }
    /// <summary>Closest standard PDF font for replacement text.</summary>
    public PdfStandardFont SuggestedFont { get; }
    /// <summary>Original PDF base-font name when available.</summary>
    public string? SourceFont { get; }
    /// <summary>Dominant font size in points.</summary>
    public double FontSize { get; }
    /// <summary>Dominant text color.</summary>
    public PdfColor Color { get; }
    /// <summary>Dominant baseline rotation in degrees.</summary>
    public double RotationDegrees { get; }
    /// <summary>Baseline X coordinate used for style-preserving replacement.</summary>
    public double BaselineX { get; }
    /// <summary>Baseline Y coordinate used for style-preserving replacement.</summary>
    public double BaselineY { get; }
}

/// <summary>Controls document text searching.</summary>
public sealed class PdfTextSearchOptions {
    private int[]? _pageNumbers;

    /// <summary>Uses ordinal case-sensitive matching when true.</summary>
    public bool MatchCase { get; set; }
    /// <summary>Requires complete word boundaries around matches when true.</summary>
    public bool WholeWords { get; set; }
    /// <summary>Includes non-Type3 invisible OCR-style text painted with PDF rendering mode 3.</summary>
    public bool IncludeTextRenderingMode3 { get; set; }
    /// <summary>Optional one-based pages to search. Null or empty searches every page.</summary>
    public int[]? PageNumbers {
        get => _pageNumbers is null ? null : (int[])_pageNumbers.Clone();
        set {
            if (value != null && value.Any(static page => page < 1)) throw new ArgumentOutOfRangeException(nameof(PageNumbers), "Search page numbers must be greater than zero.");
            if (value != null && value.Distinct().Count() != value.Length) throw new ArgumentException("Duplicate search pages are not supported.", nameof(PageNumbers));
            _pageNumbers = value is null ? null : (int[])value.Clone();
        }
    }

    internal PdfTextSearchOptions Snapshot() => new PdfTextSearchOptions { MatchCase = MatchCase, WholeWords = WholeWords, IncludeTextRenderingMode3 = IncludeTextRenderingMode3, PageNumbers = PageNumbers };
}

/// <summary>One located text occurrence in PDF user space.</summary>
public sealed class PdfTextMatch {
    internal PdfTextMatch(int pageNumber, string text, double x, double y, double width, double height, double fontSize, PdfStandardFont suggestedFont, string? sourceFont, PdfColor color, double rotationDegrees, bool usesTextRenderingMode3 = false) {
        PageNumber = pageNumber; Text = text; X = x; Y = y; Width = width; Height = height; FontSize = fontSize; SuggestedFont = suggestedFont; SourceFont = sourceFont; Color = color; RotationDegrees = rotationDegrees;
        IsTextRenderingMode3 = usesTextRenderingMode3;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Matched source text.</summary>
    public string Text { get; }
    /// <summary>Left coordinate in PDF points.</summary>
    public double X { get; }
    /// <summary>Bottom coordinate in PDF points.</summary>
    public double Y { get; }
    /// <summary>Width in PDF points.</summary>
    public double Width { get; }
    /// <summary>Height in PDF points.</summary>
    public double Height { get; }
    /// <summary>Source font size.</summary>
    public double FontSize { get; }
    /// <summary>Closest standard PDF font for replacement text.</summary>
    public PdfStandardFont SuggestedFont { get; }
    /// <summary>Original PDF base-font name when available.</summary>
    public string? SourceFont { get; }
    /// <summary>Source text color.</summary>
    public PdfColor Color { get; }
    /// <summary>Source baseline rotation.</summary>
    public double RotationDegrees { get; }
    /// <summary>True when the occurrence includes OCR-style text painted with rendering mode 3.</summary>
    public bool IsTextRenderingMode3 { get; }
}

/// <summary>Result of an existing-page text edit.</summary>
public sealed class PdfTextEditResult {
    internal PdfTextEditResult(PdfDocument document, int affectedCount, IEnumerable<string>? warnings = null) {
        Document = document;
        AffectedCount = affectedCount;
        Warnings = new ReadOnlyCollection<string>((warnings ?? Array.Empty<string>()).Distinct(StringComparer.Ordinal).ToArray());
    }

    /// <summary>Edited immutable document.</summary>
    public PdfDocument Document { get; }
    /// <summary>Number of source spans or search occurrences affected.</summary>
    public int AffectedCount { get; }
    /// <summary>Compatibility or substitution warnings.</summary>
    public IReadOnlyList<string> Warnings { get; }
}
