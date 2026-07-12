namespace OfficeIMO.Pdf;

/// <summary>
/// Line-level text block extracted from a PDF page.
/// </summary>
public sealed class PdfLogicalTextBlock : IPdfLogicalElement {
    internal PdfLogicalTextBlock(int pageNumber, PdfLogicalElementKind kind, string text, double xStart, double xEnd, double baselineY, double fontSize, int spanCount) {
        PageNumber = pageNumber;
        Kind = kind;
        Text = text;
        XStart = xStart;
        XEnd = xEnd;
        BaselineY = baselineY;
        FontSize = fontSize;
        SpanCount = spanCount;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind { get; }

    /// <summary>Extracted text for the line-level block.</summary>
    public string Text { get; }

    /// <summary>Leftmost X coordinate in PDF points.</summary>
    public double XStart { get; }

    /// <summary>Rightmost X coordinate in PDF points.</summary>
    public double XEnd { get; }

    /// <summary>Baseline Y coordinate in PDF points from the bottom of the page.</summary>
    public double BaselineY { get; }

    /// <summary>Largest font size represented by this line-level block.</summary>
    public double FontSize { get; }

    /// <summary>Number of text spans merged into this block.</summary>
    public int SpanCount { get; }
}

/// <summary>
/// Heuristic heading line inferred from text size and geometry.
/// </summary>
public sealed class PdfLogicalHeading {
    internal PdfLogicalHeading(int pageNumber, int level, string text, double fontSize, PdfLogicalTextBlock line) {
        PageNumber = pageNumber;
        Level = level;
        Text = text;
        FontSize = fontSize;
        Line = line;
        Confidence = 0.82D;
        Evidence = new[] { new PdfInferenceEvidence("heading.font-tier", "The line was assigned to a larger-font heading tier relative to nearby body text.", 0.8D) };
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Best-effort heading level, where 1 is the largest heading tier.</summary>
    public int Level { get; }

    /// <summary>Heading text.</summary>
    public string Text { get; }

    /// <summary>Representative font size in points.</summary>
    public double FontSize { get; }

    /// <summary>Line-level text block that produced the heading.</summary>
    public PdfLogicalTextBlock Line { get; }
    /// <summary>Normalized heading-classification confidence.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting the heading classification.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

/// <summary>
/// Detected bullet or numbered list item.
/// </summary>
public sealed class PdfLogicalListItem {
    internal PdfLogicalListItem(int pageNumber, int level, string marker, string text, PdfLogicalTextBlock line) {
        PageNumber = pageNumber;
        Level = level;
        Marker = marker;
        Text = text;
        Line = line;
        Confidence = string.IsNullOrWhiteSpace(marker) ? 0.55D : 0.9D;
        Evidence = new[] { new PdfInferenceEvidence(string.IsNullOrWhiteSpace(marker) ? "list.indentation" : "list.marker", string.IsNullOrWhiteSpace(marker) ? "List membership was inferred from indentation and neighboring items." : "The line begins with a recognized list marker: " + marker + ".", string.IsNullOrWhiteSpace(marker) ? 0.3D : 0.9D) };
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Best-effort nesting level, where 1 is the outermost list level.</summary>
    public int Level { get; }

    /// <summary>List marker such as "1", "1.2", "-", "•", or "(a)".</summary>
    public string Marker { get; }

    /// <summary>List item text without the marker.</summary>
    public string Text { get; }

    /// <summary>Line-level text block that produced the list item.</summary>
    public PdfLogicalTextBlock Line { get; }
    /// <summary>Normalized list-classification confidence.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting the list classification.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

/// <summary>
/// Heuristic paragraph group built from nearby line-level text blocks.
/// </summary>
public sealed class PdfLogicalParagraph {
    private PdfLogicalParagraph(
        int pageNumber,
        string text,
        IReadOnlyList<PdfLogicalTextBlock> lines,
        double xStart,
        double xEnd,
        double yTop,
        double yBottom) {
        PageNumber = pageNumber;
        Text = text;
        Lines = lines;
        XStart = xStart;
        XEnd = xEnd;
        YTop = yTop;
        YBottom = yBottom;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Paragraph text with grouped lines joined by spaces.</summary>
    public string Text { get; }

    /// <summary>Line-level blocks that make up this paragraph.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> Lines { get; }

    /// <summary>Leftmost X coordinate in PDF points.</summary>
    public double XStart { get; }

    /// <summary>Rightmost X coordinate in PDF points.</summary>
    public double XEnd { get; }

    /// <summary>Top baseline Y coordinate in PDF points.</summary>
    public double YTop { get; }

    /// <summary>Bottom baseline Y coordinate in PDF points.</summary>
    public double YBottom { get; }

    internal static PdfLogicalParagraph From(int pageNumber, StructuredParagraph paragraph, IReadOnlyList<PdfLogicalTextBlock> lines) {
        return new PdfLogicalParagraph(
            pageNumber,
            paragraph.Text,
            lines.ToArray(),
            paragraph.XStart,
            paragraph.XEnd,
            paragraph.YTop,
            paragraph.YBottom);
    }
}
