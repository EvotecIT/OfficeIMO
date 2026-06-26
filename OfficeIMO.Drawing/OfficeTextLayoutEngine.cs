using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free text layout helpers for OfficeIMO renderers.
/// </summary>
public static partial class OfficeTextLayoutEngine {
    /// <summary>
    /// Wraps text into measured lines using the supplied measurement delegate.
    /// </summary>
    /// <param name="text">Text to wrap. CRLF and CR line breaks are normalized to LF.</param>
    /// <param name="fontSize">Font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum line width before wrapping or breaking long words.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <returns>Measured wrapped lines. Empty input returns one empty line.</returns>
    public static IReadOnlyList<OfficeTextLine> WrapLines(string? text, double fontSize, double maxWidth, Func<string?, double, double> measure) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string value = text ?? string.Empty;
        double width = Math.Max(0D, maxWidth);
        string[] sourceLines = value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var output = new List<OfficeTextLine>();
        foreach (string sourceLine in sourceLines) {
            string line = sourceLine;
            if (line.Length == 0 || IsWhitespaceRun(line)) {
                output.Add(new OfficeTextLine(string.Empty, 0D));
                continue;
            }

            string current = string.Empty;
            foreach (string token in TokenizeWhitespaceRuns(line)) {
                bool whitespace = IsWhitespaceRun(token);
                if (!whitespace && current.Length == 0 && Measure(token, fontSize, measure) > width) {
                    foreach (OfficeTextLine part in BreakWord(token, fontSize, width, measure)) {
                        output.Add(part);
                    }

                    continue;
                }

                string candidate = current + token;
                if (current.Length > 0 && Measure(candidate, fontSize, measure) > width) {
                    string emitted = TrimTrailingSoftWrapWhitespace(current);
                    if (emitted.Length > 0) {
                        output.Add(new OfficeTextLine(emitted, Measure(emitted, fontSize, measure)));
                    }

                    current = whitespace ? string.Empty : token;
                    if (!whitespace && Measure(current, fontSize, measure) > width) {
                        foreach (OfficeTextLine part in BreakWord(current, fontSize, width, measure)) {
                            output.Add(part);
                        }

                        current = string.Empty;
                    }
                } else {
                    current = candidate;
                }
            }

            if (current.Length > 0) {
                string emitted = TrimTrailingSoftWrapWhitespace(current);
                if (emitted.Length > 0) {
                    output.Add(new OfficeTextLine(emitted, Measure(emitted, fontSize, measure)));
                }
            }
        }

        return output.Count == 0
            ? new[] { new OfficeTextLine(string.Empty, 0D) }
            : output;
    }

    private static IEnumerable<string> TokenizeWhitespaceRuns(string text) {
        int start = 0;
        bool whitespace = IsSoftWrapWhitespace(text[0]);
        for (int i = 1; i < text.Length; i++) {
            bool currentWhitespace = IsSoftWrapWhitespace(text[i]);
            if (currentWhitespace != whitespace) {
                yield return text.Substring(start, i - start);
                start = i;
                whitespace = currentWhitespace;
            }
        }

        yield return text.Substring(start);
    }

    private static bool IsWhitespaceRun(string text) {
        for (int i = 0; i < text.Length; i++) {
            if (!IsSoftWrapWhitespace(text[i])) {
                return false;
            }
        }

        return text.Length > 0;
    }

    private static bool IsSoftWrapWhitespace(char value) =>
        value == ' ' || value == '\t';

    private static string TrimTrailingSoftWrapWhitespace(string text) {
        int end = text.Length;
        while (end > 0 && IsSoftWrapWhitespace(text[end - 1])) {
            end--;
        }

        return end == text.Length ? text : text.Substring(0, end);
    }

    /// <summary>
    /// Gets the widest measured line.
    /// </summary>
    /// <param name="lines">Measured text lines.</param>
    /// <returns>The maximum <see cref="OfficeTextLine.Width"/> value, or zero for an empty set.</returns>
    public static double MeasureMaxLineWidth(IReadOnlyList<OfficeTextLine> lines) {
        if (lines == null) {
            throw new ArgumentNullException(nameof(lines));
        }

        double max = 0D;
        for (int i = 0; i < lines.Count; i++) {
            max = Math.Max(max, lines[i].Width);
        }

        return max;
    }

    /// <summary>
    /// Measures a single line and trims it with an ellipsis when it exceeds the requested width.
    /// </summary>
    /// <param name="text">Text to measure and trim.</param>
    /// <param name="fontSize">Font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum accepted line width.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="clipped">Set to <c>true</c> when the returned line had to be shortened.</param>
    /// <returns>The measured line, possibly shortened with an ellipsis.</returns>
    public static OfficeTextLine TrimLineToWidth(string? text, double fontSize, double maxWidth, Func<string?, double, double> measure, out bool clipped) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string value = text ?? string.Empty;
        double width = Math.Max(0D, maxWidth);
        double measured = Measure(value, fontSize, measure);
        if (measured <= width) {
            clipped = false;
            return new OfficeTextLine(value, measured);
        }

        clipped = true;
        const string ellipsis = "...";
        while (value.Length > 0 && Measure(value + ellipsis, fontSize, measure) > width) {
            value = OfficeTextElements.RemoveLast(value);
        }

        value = value.Length == 0 && Measure(ellipsis, fontSize, measure) > width ? string.Empty : value + ellipsis;
        return new OfficeTextLine(value, Measure(value, fontSize, measure));
    }

    /// <summary>
    /// Measures a single line and trims the beginning with an ellipsis when it exceeds the requested width.
    /// </summary>
    /// <param name="text">Text to measure and trim.</param>
    /// <param name="fontSize">Font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum accepted line width.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="clipped">Set to <c>true</c> when the returned line had to be shortened.</param>
    /// <returns>The measured line, possibly shortened with a leading ellipsis.</returns>
    public static OfficeTextLine TrimLineStartToWidth(string? text, double fontSize, double maxWidth, Func<string?, double, double> measure, out bool clipped) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string value = text ?? string.Empty;
        double width = Math.Max(0D, maxWidth);
        double measured = Measure(value, fontSize, measure);
        if (measured <= width) {
            clipped = false;
            return new OfficeTextLine(value, measured);
        }

        clipped = true;
        const string ellipsis = "...";
        while (value.Length > 0 && Measure(ellipsis + value, fontSize, measure) > width) {
            value = OfficeTextElements.RemoveFirst(value);
        }

        value = value.Length == 0 && Measure(ellipsis, fontSize, measure) > width ? string.Empty : ellipsis + value;
        return new OfficeTextLine(value, Measure(value, fontSize, measure));
    }

    /// <summary>
    /// Finds the largest single-line font size that fits within the requested width.
    /// </summary>
    /// <param name="text">Text to measure.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum accepted line width.</param>
    /// <param name="minimumFontSize">Minimum font size when fitting is required.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="iterations">Number of binary-search iterations used between the minimum and initial font size.</param>
    /// <returns>The initial font size when text already fits; otherwise the largest measured size above the minimum that fits.</returns>
    public static double FitSingleLineFontSize(
        string? text,
        double fontSize,
        double maxWidth,
        double minimumFontSize,
        Func<string?, double, double> measure,
        int iterations = 10) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedFontSize = NormalizePositive(fontSize, 1D);
        double minFontSize = Math.Min(resolvedFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
        double width = NormalizeNonNegative(maxWidth);
        if (string.IsNullOrEmpty(text) || Measure(text, resolvedFontSize, measure) <= width) {
            return resolvedFontSize;
        }

        if (Measure(text, minFontSize, measure) > width) {
            return minFontSize;
        }

        double low = minFontSize;
        double high = resolvedFontSize;
        int count = Math.Max(1, iterations);
        for (int i = 0; i < count; i++) {
            double candidate = (low + high) / 2D;
            if (Measure(text, candidate, measure) <= width) {
                low = candidate;
            } else {
                high = candidate;
            }
        }

        return low;
    }

    /// <summary>
    /// Estimates the maximum unrotated single-line text width that can remain inside a rotated bounding rectangle.
    /// </summary>
    /// <param name="availableWidth">Available unrotated rectangle width.</param>
    /// <param name="availableHeight">Available unrotated rectangle height.</param>
    /// <param name="lineHeight">Estimated rendered line height.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <returns>A positive width limit that callers can pass to single-line rotated text layout.</returns>
    public static double ResolveRotatedTextWidthLimit(double availableWidth, double availableHeight, double lineHeight, double rotationDegrees) {
        double width = Math.Max(1D, NormalizeNonNegative(availableWidth));
        double height = Math.Max(1D, NormalizeNonNegative(availableHeight));
        double radians = Math.Abs(rotationDegrees) * Math.PI / 180D;
        double cos = Math.Abs(Math.Cos(radians));
        double sin = Math.Abs(Math.Sin(radians));
        double estimatedHeight = Math.Max(1D, NormalizePositive(lineHeight, 1D));
        double limit = Math.Max(width, height);

        if (cos > 0.000001D) {
            limit = Math.Min(limit, (width - (estimatedHeight * sin)) / cos);
        }

        if (sin > 0.000001D) {
            limit = Math.Min(limit, (height - (estimatedHeight * cos)) / sin);
        }

        if (double.IsNaN(limit) || double.IsInfinity(limit)) {
            return Math.Max(width, height);
        }

        return Math.Max(1D, limit);
    }

    /// <summary>
    /// Lays out a bounded text block with optional wrapping, single-line normalization, shrink-to-fit, and height clipping.
    /// </summary>
    /// <param name="text">Text to lay out.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used to derive line height from font size.</param>
    /// <param name="minimumFontSize">Minimum font size when single-line shrink-to-fit is enabled.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled.</param>
    /// <param name="forceSingleLine">Whether line breaks should be normalized to spaces and wrapping disabled.</param>
    /// <param name="shrinkToFit">Whether single-line text should reduce font size to fit the requested width.</param>
    /// <returns>Measured text block with the resolved font size, line height, width, height, lines, and clipping state.</returns>
    public static OfficeTextBlockLayout LayoutTextBlock(
        string? text,
        double fontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        double minimumFontSize,
        Func<string?, double, double> measure,
        bool wrap,
        bool forceSingleLine = false,
        bool shrinkToFit = false) =>
        LayoutTextBlock(
            text,
            fontSize,
            maxWidth,
            maxHeight,
            lineHeightFactor,
            minimumFontSize,
            measure,
            wrap,
            forceSingleLine,
            shrinkToFit,
            OfficeTextOverflowBehavior.Ellipsis);

    /// <summary>
    /// Lays out a bounded text block with optional wrapping, single-line normalization, shrink-to-fit, overflow policy, and height clipping.
    /// </summary>
    /// <param name="text">Text to lay out.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used to derive line height from font size.</param>
    /// <param name="minimumFontSize">Minimum font size when single-line shrink-to-fit is enabled.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled.</param>
    /// <param name="forceSingleLine">Whether line breaks should be normalized to spaces and wrapping disabled.</param>
    /// <param name="shrinkToFit">Whether single-line text should reduce font size to fit the requested width.</param>
    /// <param name="overflowBehavior">How overflowing text should be represented in the returned layout.</param>
    /// <returns>Measured text block with the resolved font size, line height, width, height, lines, and clipping state.</returns>
    public static OfficeTextBlockLayout LayoutTextBlock(
        string? text,
        double fontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        double minimumFontSize,
        Func<string?, double, double> measure,
        bool wrap,
        bool forceSingleLine,
        bool shrinkToFit,
        OfficeTextOverflowBehavior overflowBehavior) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string layoutText = forceSingleLine ? NormalizeSingleLineText(text ?? string.Empty) : text ?? string.Empty;
        bool hasHardBreaks = !forceSingleLine && (layoutText.IndexOf('\n') >= 0 || layoutText.IndexOf('\r') >= 0);
        bool effectiveWrap = !forceSingleLine && (wrap || hasHardBreaks);
        double resolvedFontSize = NormalizePositive(fontSize, 1D);
        double minFontSize = Math.Min(resolvedFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        double layoutFontSize = shrinkToFit && !effectiveWrap
            ? FitSingleLineFontSize(layoutText, resolvedFontSize, width, minFontSize, measure)
            : resolvedFontSize;
        double lineHeight = Math.Max(1D, Math.Ceiling(layoutFontSize * lineFactor));
        IReadOnlyList<OfficeTextLine> lines;
        bool clipped = false;

        if (effectiveWrap) {
            lines = WrapLines(layoutText, layoutFontSize, width, measure);
        } else {
            string normalized = layoutText.Replace("\r\n", "\n").Replace('\r', '\n');
            string firstLine = normalized.Split('\n')[0];
            OfficeTextLine line = ResolveOverflowLine(firstLine, layoutFontSize, width, measure, overflowBehavior, out bool lineClipped);
            clipped = lineClipped;
            lines = new[] { line };
        }

        return ClipTextBlockToHeight(lines, layoutFontSize, lineHeight, width, height, measure, clipped, overflowBehavior);
    }

    /// <summary>
    /// Wraps and measures a text block, reducing font size when the measured block does not fit the requested bounds.
    /// </summary>
    /// <param name="text">Text to wrap and measure.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used to derive line height from font size.</param>
    /// <param name="minimumFontSize">Minimum font size when scaling down to fit.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <returns>Measured text block with the resolved font size, line height, width, height, and lines.</returns>
    public static OfficeTextBlockLayout FitWrappedText(
        string? text,
        double fontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        double minimumFontSize,
        Func<string?, double, double> measure) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedFontSize = NormalizePositive(fontSize, 1D);
        double minFontSize = Math.Max(1D, NormalizePositive(minimumFontSize, 1D));
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        OfficeTextBlockLayout layout = CreateBlockLayout(text, resolvedFontSize, width, lineFactor, measure);
        double scaleDown = Math.Min(1D, Math.Min(width / Math.Max(layout.Width, 1D), height / Math.Max(layout.Height, 1D)));
        if (scaleDown < 0.98D) {
            resolvedFontSize = Math.Max(minFontSize, resolvedFontSize * Math.Max(0D, scaleDown));
            layout = CreateBlockLayout(text, resolvedFontSize, width, lineFactor, measure);
        }

        return layout;
    }

    /// <summary>
    /// Clips measured text lines to the requested block height and ellipsizes the last visible line when lines are omitted.
    /// </summary>
    /// <param name="lines">Measured text lines to clip.</param>
    /// <param name="fontSize">Resolved font size used for measurement.</param>
    /// <param name="lineHeight">Resolved line height.</param>
    /// <param name="maxWidth">Maximum line width used for ellipsis trimming.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="alreadyClipped">Whether an earlier layout stage already clipped or ellipsized the text.</param>
    /// <returns>A measured text block whose visible lines fit the requested height.</returns>
    public static OfficeTextBlockLayout ClipTextBlockToHeight(
        IReadOnlyList<OfficeTextLine> lines,
        double fontSize,
        double lineHeight,
        double maxWidth,
        double maxHeight,
        Func<string?, double, double> measure,
        bool alreadyClipped = false) =>
        ClipTextBlockToHeight(
            lines,
            fontSize,
            lineHeight,
            maxWidth,
            maxHeight,
            measure,
            alreadyClipped,
            OfficeTextOverflowBehavior.Ellipsis);

    /// <summary>
    /// Clips measured text lines to the requested block height and applies the requested overflow policy to omitted lines.
    /// </summary>
    /// <param name="lines">Measured text lines to clip.</param>
    /// <param name="fontSize">Resolved font size used for measurement.</param>
    /// <param name="lineHeight">Resolved line height.</param>
    /// <param name="maxWidth">Maximum line width used for ellipsis trimming.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="alreadyClipped">Whether an earlier layout stage already clipped or ellipsized the text.</param>
    /// <param name="overflowBehavior">How omitted or oversized text should be represented in the returned layout.</param>
    /// <returns>A measured text block whose visible lines fit the requested height.</returns>
    public static OfficeTextBlockLayout ClipTextBlockToHeight(
        IReadOnlyList<OfficeTextLine> lines,
        double fontSize,
        double lineHeight,
        double maxWidth,
        double maxHeight,
        Func<string?, double, double> measure,
        bool alreadyClipped,
        OfficeTextOverflowBehavior overflowBehavior) {
        if (lines == null) {
            throw new ArgumentNullException(nameof(lines));
        }

        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedFontSize = NormalizePositive(fontSize, 1D);
        double resolvedLineHeight = NormalizePositive(lineHeight, resolvedFontSize);
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        int maxLines = Math.Max(1, (int)Math.Floor(height / resolvedLineHeight));
        bool clipped = alreadyClipped;
        var visible = new List<OfficeTextLine>(Math.Min(lines.Count, maxLines));
        int count = Math.Min(lines.Count, maxLines);
        for (int i = 0; i < count; i++) {
            visible.Add(lines[i]);
        }

        if (lines.Count > maxLines) {
            clipped = true;
            if (visible.Count > 0) {
                OfficeTextLine last = visible[visible.Count - 1];
                if (overflowBehavior == OfficeTextOverflowBehavior.Ellipsis) {
                    visible[visible.Count - 1] = TrimLineToWidth(last.Text + "...", resolvedFontSize, width, measure, out _);
                }
            }
        }

        if (visible.Count == 0) {
            visible.Add(new OfficeTextLine(string.Empty, 0D));
        }

        double blockWidth = MeasureMaxLineWidth(visible);
        double blockHeight = visible.Count * resolvedLineHeight;
        return new OfficeTextBlockLayout(visible, resolvedFontSize, resolvedLineHeight, blockWidth, blockHeight, clipped);
    }

    private static IEnumerable<OfficeTextLine> BreakWord(string word, double fontSize, double maxWidth, Func<string?, double, double> measure) {
        string part = string.Empty;
        foreach (string textElement in OfficeTextElements.Enumerate(word)) {
            string candidate = part + textElement;
            if (part.Length > 0 && Measure(candidate, fontSize, measure) > maxWidth) {
                yield return new OfficeTextLine(part, Measure(part, fontSize, measure));
                part = string.Empty;
            }

            part += textElement;
        }

        if (part.Length > 0) {
            yield return new OfficeTextLine(part, Measure(part, fontSize, measure));
        }
    }

    private static OfficeTextLine ResolveOverflowLine(
        string? text,
        double fontSize,
        double maxWidth,
        Func<string?, double, double> measure,
        OfficeTextOverflowBehavior overflowBehavior,
        out bool clipped) {
        string value = text ?? string.Empty;
        double width = Math.Max(0D, maxWidth);
        double measured = Measure(value, fontSize, measure);
        if (measured <= width || overflowBehavior != OfficeTextOverflowBehavior.Clip) {
            return TrimLineToWidth(value, fontSize, width, measure, out clipped);
        }

        clipped = true;
        return new OfficeTextLine(value, measured);
    }

    private static double Measure(string? text, double fontSize, Func<string?, double, double> measure) =>
        string.IsNullOrEmpty(text) ? 0D : measure(text, fontSize);

    private static string NormalizeSingleLineText(string text) =>
        text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');

    private static OfficeTextBlockLayout CreateBlockLayout(string? text, double fontSize, double maxWidth, double lineHeightFactor, Func<string?, double, double> measure) {
        IReadOnlyList<OfficeTextLine> lines = WrapLines(text, fontSize, maxWidth, measure);
        double lineHeight = fontSize * lineHeightFactor;
        double width = MeasureMaxLineWidth(lines);
        double height = Math.Max(fontSize, ((lines.Count - 1) * lineHeight) + fontSize);
        return new OfficeTextBlockLayout(lines, fontSize, lineHeight, width, height);
    }

    private static double NormalizePositive(double value, double fallback) =>
        value > 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && !double.IsNaN(value) ? value : 0D;
}
