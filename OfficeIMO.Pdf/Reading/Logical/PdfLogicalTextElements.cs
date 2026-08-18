namespace OfficeIMO.Pdf;

/// <summary>
/// Line-level text block extracted from a PDF page.
/// </summary>
public sealed class PdfLogicalTextBlock : IPdfLogicalElement {
    internal PdfLogicalTextBlock(
        int pageNumber,
        PdfLogicalElementKind kind,
        string text,
        double xStart,
        double xEnd,
        double baselineY,
        double fontSize,
        IReadOnlyList<PdfTextSpan> spans,
        PdfLogicalContentSourceKind sourceKind = PdfLogicalContentSourceKind.Native,
        double confidence = 1D,
        PdfLogicalVisualBounds? visualBounds = null) {
        PageNumber = pageNumber;
        Kind = kind;
        Text = text;
        XStart = xStart;
        XEnd = xEnd;
        BaselineY = baselineY;
        FontSize = fontSize;
        Spans = Array.AsReadOnly((spans ?? throw new ArgumentNullException(nameof(spans))).ToArray());
        Runs = BuildRuns(text, Spans);
        SourceKind = sourceKind;
        Confidence = PdfInference.Clamp(confidence);
        VisualBounds = visualBounds;
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

    /// <summary>Positioned text spans merged into this block, preserving font, color, and geometry details.</summary>
    public IReadOnlyList<PdfTextSpan> Spans { get; }

    /// <summary>
    /// Line text segmented into reusable semantic runs aligned with the source span styles.
    /// Whitespace synthesized by layout analysis is retained in the adjacent run.
    /// </summary>
    public IReadOnlyList<PdfLogicalTextRun> Runs { get; }

    /// <summary>Whether the block came from native PDF text or an external OCR provider.</summary>
    public PdfLogicalContentSourceKind SourceKind { get; }

    /// <summary>Normalized extraction or provider confidence from 0 through 1.</summary>
    public double Confidence { get; }

    /// <summary>
    /// Direct top-left visual geometry when the source already supplied normalized page coordinates.
    /// Native text normally derives this geometry from PDF user-space spans on demand.
    /// </summary>
    public PdfLogicalVisualBounds? VisualBounds { get; }

    /// <summary>Number of text spans merged into this block.</summary>
    public int SpanCount => Spans.Count;

    private static IReadOnlyList<PdfLogicalTextRun> BuildRuns(
        string text,
        IReadOnlyList<PdfTextSpan> spans) {
        if (string.IsNullOrEmpty(text)) {
            return Array.Empty<PdfLogicalTextRun>();
        }

        var runs = new List<PdfLogicalTextRun>();
        int cursor = 0;
        for (int i = 0; i < spans.Count && cursor < text.Length; i++) {
            PdfTextSpan span = spans[i];
            if (string.IsNullOrEmpty(span.Text)) {
                continue;
            }

            if (!TryFindNormalizedSpan(text, span.Text, cursor, out int index, out int end)) {
                continue;
            }

            string runText = text.Substring(cursor, end - cursor);
            if (runs.Count > 0 && StylesMatch(runs[runs.Count - 1].SourceSpan, span)) {
                PdfLogicalTextRun previous = runs[runs.Count - 1];
                runs[runs.Count - 1] = new PdfLogicalTextRun(previous.Text + runText, previous.SourceSpan);
            } else {
                runs.Add(new PdfLogicalTextRun(runText, span));
            }
            cursor = end;
        }

        if (runs.Count == 0) {
            return new[] { new PdfLogicalTextRun(text, sourceSpan: null) };
        }

        if (cursor < text.Length) {
            PdfLogicalTextRun last = runs[runs.Count - 1];
            var combined = new System.Text.StringBuilder(last.Text.Length + text.Length - cursor);
            combined.Append(last.Text);
            combined.Append(text, cursor, text.Length - cursor);
            runs[runs.Count - 1] = new PdfLogicalTextRun(
                combined.ToString(),
                last.SourceSpan);
        }

        return Array.AsReadOnly(runs.ToArray());
    }

    private static bool TryFindNormalizedSpan(
        string text,
        string spanText,
        int start,
        out int index,
        out int end) {
        index = -1;
        end = -1;

        int sourceStart = 0;
        while (sourceStart < spanText.Length && char.IsWhiteSpace(spanText[sourceStart])) {
            sourceStart++;
        }

        if (sourceStart == spanText.Length) {
            return false;
        }

        for (int candidate = start; candidate < text.Length; candidate++) {
            if (text[candidate] != spanText[sourceStart]) {
                continue;
            }

            int source = sourceStart;
            int target = candidate;
            while (source < spanText.Length) {
                if (char.IsWhiteSpace(spanText[source])) {
                    while (source < spanText.Length && char.IsWhiteSpace(spanText[source])) {
                        source++;
                    }
                    while (target < text.Length && char.IsWhiteSpace(text[target])) {
                        target++;
                    }
                    continue;
                }

                if (target >= text.Length || text[target] != spanText[source]) {
                    break;
                }

                source++;
                target++;
            }

            if (source == spanText.Length) {
                index = candidate;
                end = target;
                return true;
            }
        }

        return false;
    }

    private static bool StylesMatch(PdfTextSpan? left, PdfTextSpan right) {
        if (left == null) {
            return false;
        }

        return string.Equals(left.BaseFont, right.BaseFont, StringComparison.Ordinal) &&
            Math.Abs(left.FontSize - right.FontSize) <= 0.001D &&
            left.Color == right.Color &&
            left.IsVisible == right.IsVisible;
    }
}

/// <summary>
/// Semantic text fragment aligned with one positioned PDF text span.
/// </summary>
public sealed class PdfLogicalTextRun {
    internal PdfLogicalTextRun(string text, PdfTextSpan? sourceSpan) {
        Text = text ?? string.Empty;
        SourceSpan = sourceSpan;
    }

    /// <summary>Text in this semantic fragment, including inferred adjacent whitespace.</summary>
    public string Text { get; }

    /// <summary>Positioned source span supplying style and geometry, when alignment succeeded.</summary>
    public PdfTextSpan? SourceSpan { get; }

    /// <summary>Source PDF base font name, when available.</summary>
    public string? BaseFont => SourceSpan?.BaseFont;

    /// <summary>Source font size in points, or zero when no source span was aligned.</summary>
    public double FontSize => SourceSpan?.FontSize ?? 0D;

    /// <summary>Source text color, when available.</summary>
    public OfficeIMO.Drawing.OfficeColor? Color => SourceSpan?.Color;

    /// <summary>Best-effort bold classification derived from the source PDF base font name.</summary>
    public bool IsBold => HasFontStyle("Bold") || HasFontStyle("Black") || HasFontStyle("Demi");

    /// <summary>Best-effort italic classification derived from the source PDF base font name.</summary>
    public bool IsItalic => HasFontStyle("Italic") || HasFontStyle("Oblique");

    private bool HasFontStyle(string token) =>
        BaseFont?.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
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
        Runs = SliceRuns(line, text);
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

    /// <summary>List item text segmented into semantic runs, excluding the detected marker.</summary>
    public IReadOnlyList<PdfLogicalTextRun> Runs { get; }

    /// <summary>Normalized list-classification confidence.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting the list classification.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }

    private static IReadOnlyList<PdfLogicalTextRun> SliceRuns(PdfLogicalTextBlock line, string text) {
        if (string.IsNullOrEmpty(text)) {
            return Array.Empty<PdfLogicalTextRun>();
        }

        int start = line.Text.IndexOf(text, StringComparison.Ordinal);
        if (start < 0 || line.Runs.Count == 0) {
            return new[] { new PdfLogicalTextRun(text, sourceSpan: null) };
        }

        int end = start + text.Length;
        int cursor = 0;
        var result = new List<PdfLogicalTextRun>();
        for (int i = 0; i < line.Runs.Count && cursor < end; i++) {
            PdfLogicalTextRun run = line.Runs[i];
            int runStart = cursor;
            int runEnd = cursor + run.Text.Length;
            cursor = runEnd;
            int sliceStart = Math.Max(start, runStart);
            int sliceEnd = Math.Min(end, runEnd);
            if (sliceStart >= sliceEnd) {
                continue;
            }

            result.Add(new PdfLogicalTextRun(
                run.Text.Substring(sliceStart - runStart, sliceEnd - sliceStart),
                run.SourceSpan));
        }

        return result.Count == 0
            ? new[] { new PdfLogicalTextRun(text, sourceSpan: null) }
            : Array.AsReadOnly(result.ToArray());
    }
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

    internal static PdfLogicalParagraph FromOcr(int pageNumber, PdfLogicalTextBlock line) {
        Guard.NotNull(line, nameof(line));
        return FromOcr(pageNumber, new[] { line });
    }

    internal static PdfLogicalParagraph FromOcr(int pageNumber, IReadOnlyList<PdfLogicalTextBlock> lines) {
        Guard.NotNull(lines, nameof(lines));
        if (lines.Count == 0) throw new ArgumentException("At least one OCR line is required.", nameof(lines));
        return new PdfLogicalParagraph(
            pageNumber,
            string.Join(" ", lines.Select(static line => line.Text)),
            lines.ToArray(),
            lines.Min(static line => line.XStart),
            lines.Max(static line => line.XEnd),
            lines.Max(static line => line.BaselineY),
            lines.Min(static line => line.BaselineY));
    }
}
