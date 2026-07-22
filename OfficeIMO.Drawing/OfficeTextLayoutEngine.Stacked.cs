using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    /// <summary>
    /// Lays out text as upright stacked text elements for vertical cell and shape renderers.
    /// </summary>
    /// <param name="text">Text to stack. CR/LF breaks are ignored because each text element becomes its own line.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum stacked block width.</param>
    /// <param name="maxHeight">Maximum stacked block height.</param>
    /// <param name="lineHeightFactor">Multiplier used to derive line height from font size.</param>
    /// <param name="minimumFontSize">Minimum font size when scaling down to fit.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="shrinkToFit">Whether the stacked block should reduce font size to fit both width and height.</param>
    /// <returns>A measured text block with one Unicode text element per line.</returns>
    public static OfficeTextBlockLayout LayoutStackedTextBlock(
        string? text,
        double fontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        double minimumFontSize,
        Func<string?, double, double> measure,
        bool shrinkToFit = true) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string value = NormalizeStackedText(text);
        double resolvedFontSize = NormalizePositive(fontSize, 1D);
        double minFontSize = Math.Min(resolvedFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        if (value.Length == 0) {
            return new OfficeTextBlockLayout(new[] { new OfficeTextLine(string.Empty, 0D) }, resolvedFontSize, Math.Max(1D, Math.Ceiling(resolvedFontSize * lineFactor)), 0D, Math.Max(1D, Math.Ceiling(resolvedFontSize * lineFactor)));
        }

        IReadOnlyList<string> elements = SplitTextElements(value);
        if (shrinkToFit) {
            resolvedFontSize = FitStackedFontSize(elements, resolvedFontSize, minFontSize, width, height, lineFactor, measure);
        }

        double lineHeight = Math.Max(1D, Math.Ceiling(resolvedFontSize * lineFactor));
        List<OfficeTextLine> lines = CreateStackedLines(elements, resolvedFontSize, measure);
        return ClipTextBlockToHeight(lines, resolvedFontSize, lineHeight, width, height, measure);
    }

    /// <summary>
    /// Lays out styled rich text runs as upright stacked text elements for vertical cell and shape renderers.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum stacked block width.</param>
    /// <param name="maxHeight">Maximum stacked block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="shrinkToFit">Whether the stacked block should reduce font sizes to fit both width and height.</param>
    /// <param name="minimumFontSize">Minimum font size when scaling down to fit.</param>
    /// <returns>A measured rich text block with one Unicode text element per line.</returns>
    public static OfficeRichTextBlockLayout LayoutStackedRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure,
        bool shrinkToFit = true,
        double minimumFontSize = 1D) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        return LayoutStackedRichTextBlock(
            runs,
            maxWidth,
            maxHeight,
            lineHeightFactor,
            (text, fontSize, _) => measure(text, fontSize),
            shrinkToFit,
            minimumFontSize);
    }

    /// <summary>
    /// Lays out styled rich text runs as upright stacked text elements for vertical cell and shape renderers.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum stacked block width.</param>
    /// <param name="maxHeight">Maximum stacked block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double, string?)"/>.</param>
    /// <param name="shrinkToFit">Whether the stacked block should reduce font sizes to fit both width and height.</param>
    /// <param name="minimumFontSize">Minimum font size when scaling down to fit.</param>
    /// <returns>A measured rich text block with one Unicode text element per line.</returns>
    public static OfficeRichTextBlockLayout LayoutStackedRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, double> measure,
        bool shrinkToFit = true,
        double minimumFontSize = 1D) {
        if (runs == null) {
            throw new ArgumentNullException(nameof(runs));
        }

        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        IReadOnlyList<OfficeRichTextRun> elements = SplitRichTextElements(NormalizeRichTextRuns(runs));
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        if (elements.Count == 0) {
            double lineHeight = Math.Max(1D, Math.Ceiling(lineFactor));
            return new OfficeRichTextBlockLayout(new[] { new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), lineHeight) }, lineHeight, 0D, lineHeight);
        }

        if (shrinkToFit) {
            elements = FitStackedRichTextRuns(elements, width, height, lineFactor, minimumFontSize, measure);
        }

        double maxFontSize = ResolveMaxRichTextFontSize(elements);
        double resolvedLineHeight = Math.Max(1D, Math.Ceiling(maxFontSize * lineFactor));
        List<OfficeRichTextLine> lines = CreateStackedRichTextLines(elements, measure, resolvedLineHeight);
        return ClipStackedRichTextBlockToHeight(lines, resolvedLineHeight, width, height, measure);
    }

    private static double FitStackedFontSize(
        IReadOnlyList<string> elements,
        double fontSize,
        double minimumFontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure) {
        if (StackedFits(elements, fontSize, maxWidth, maxHeight, lineHeightFactor, measure)) {
            return fontSize;
        }

        if (!StackedFits(elements, minimumFontSize, maxWidth, maxHeight, lineHeightFactor, measure)) {
            return minimumFontSize;
        }

        double low = minimumFontSize;
        double high = fontSize;
        for (int i = 0; i < 10; i++) {
            double candidate = (low + high) / 2D;
            if (StackedFits(elements, candidate, maxWidth, maxHeight, lineHeightFactor, measure)) {
                low = candidate;
            } else {
                high = candidate;
            }
        }

        return low;
    }

    private static bool StackedFits(
        IReadOnlyList<string> elements,
        double fontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure) {
        double lineHeight = Math.Max(1D, Math.Ceiling(fontSize * lineHeightFactor));
        double height = elements.Count * lineHeight;
        if (height > maxHeight) {
            return false;
        }

        for (int i = 0; i < elements.Count; i++) {
            if (Measure(elements[i], fontSize, measure) > maxWidth) {
                return false;
            }
        }

        return true;
    }

    private static List<OfficeTextLine> CreateStackedLines(IReadOnlyList<string> elements, double fontSize, Func<string?, double, double> measure) {
        var lines = new List<OfficeTextLine>(elements.Count);
        for (int i = 0; i < elements.Count; i++) {
            string element = elements[i];
            lines.Add(new OfficeTextLine(element, Measure(element, fontSize, measure)));
        }

        return lines;
    }

    private static IReadOnlyList<OfficeRichTextRun> FitStackedRichTextRuns(
        IReadOnlyList<OfficeRichTextRun> elements,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        double minimumFontSize,
        Func<string?, double, string?, double> measure) {
        double fontSize = ResolveMaxRichTextFontSize(elements);
        double minFontSize = Math.Min(fontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
        if (StackedRichTextFits(elements, maxWidth, maxHeight, lineHeightFactor, measure)) {
            return elements;
        }

        double minScale = minFontSize / Math.Max(fontSize, 1D);
        IReadOnlyList<OfficeRichTextRun> minimumRuns = ScaleRichTextRuns(elements, minScale);
        if (!StackedRichTextFits(minimumRuns, maxWidth, maxHeight, lineHeightFactor, measure)) {
            return minimumRuns;
        }

        double low = minScale;
        double high = 1D;
        for (int i = 0; i < 10; i++) {
            double candidate = (low + high) / 2D;
            IReadOnlyList<OfficeRichTextRun> scaled = ScaleRichTextRuns(elements, candidate);
            if (StackedRichTextFits(scaled, maxWidth, maxHeight, lineHeightFactor, measure)) {
                low = candidate;
            } else {
                high = candidate;
            }
        }

        return ScaleRichTextRuns(elements, low);
    }

    private static bool StackedRichTextFits(
        IReadOnlyList<OfficeRichTextRun> elements,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, double> measure) {
        double fontSize = ResolveMaxRichTextFontSize(elements);
        double lineHeight = Math.Max(1D, Math.Ceiling(fontSize * lineHeightFactor));
        if (elements.Count * lineHeight > maxHeight) {
            return false;
        }

        for (int i = 0; i < elements.Count; i++) {
            OfficeRichTextRun run = elements[i];
            if (Measure(run.Text, run.FontSize, run.FontFamily, measure) > maxWidth) {
                return false;
            }
        }

        return true;
    }

    private static List<OfficeRichTextLine> CreateStackedRichTextLines(IReadOnlyList<OfficeRichTextRun> elements, Func<string?, double, string?, double> measure, double lineHeight) {
        var lines = new List<OfficeRichTextLine>(elements.Count);
        for (int i = 0; i < elements.Count; i++) {
            OfficeRichTextRun run = elements[i];
            lines.Add(new OfficeRichTextLine(new[] {
                new OfficeRichTextSegment(
                    run.Text,
                    Measure(run.Text, run.FontSize, run.FontFamily, measure),
                    run.FontSize,
                    run.Color,
                    run.Bold,
                    run.Italic,
                    run.Underline,
                    run.FontFamily,
                    run.Strikethrough,
                    run.BackgroundColor)
            }, lineHeight));
        }

        return lines;
    }

    private static OfficeRichTextBlockLayout ClipStackedRichTextBlockToHeight(
        List<OfficeRichTextLine> lines,
        double lineHeight,
        double maxWidth,
        double maxHeight,
        Func<string?, double, string?, double> measure) {
        bool clipped = false;
        int maxLines = Math.Max(1, (int)Math.Floor(NormalizeNonNegative(maxHeight) / Math.Max(1D, lineHeight)));
        if (lines.Count > maxLines) {
            clipped = true;
            lines.RemoveRange(maxLines, lines.Count - maxLines);
        }

        for (int i = 0; i < lines.Count; i++) {
            if (lines[i].Width <= maxWidth + 0.01D) {
                continue;
            }

            clipped = true;
            lines[i] = TrimRichTextLineToWidthWithEllipsis(lines[i], maxWidth, measure);
        }

        double blockWidth = MeasureMaxRichTextLineWidth(lines);
        double blockHeight = MeasureRichTextBlockHeight(lines, lineHeight);
        return new OfficeRichTextBlockLayout(lines, lineHeight, blockWidth, blockHeight, clipped);
    }

    private static IReadOnlyList<string> SplitTextElements(string text) {
        var elements = new List<string>();
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            elements.Add(element);
            if (elements.Count >= MaximumLayoutLines) break;
        }

        if (elements.Count == 0) elements.Add(string.Empty);
        return elements;
    }

    private static IReadOnlyList<OfficeRichTextRun> SplitRichTextElements(IReadOnlyList<OfficeRichTextRun> runs) {
        var elements = new List<OfficeRichTextRun>();
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i];
            string value = NormalizeStackedText(run.Text);
            if (value.Length == 0) {
                continue;
            }

            foreach (string textElement in OfficeTextElements.Enumerate(value)) {
                elements.Add(new OfficeRichTextRun(
                    textElement,
                    run.FontSize,
                    run.Color,
                    run.Bold,
                    run.Italic,
                    run.Underline,
                    run.FontFamily,
                    run.Strikethrough,
                    run.BackgroundColor));
                if (elements.Count >= MaximumLayoutLines) return elements;
            }
        }

        return elements;
    }

    private static string NormalizeStackedText(string? text) =>
        LimitLayoutText(ExpandTabs(LimitLayoutText(text ?? string.Empty)))
            .Replace("\r\n", string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty);
}
