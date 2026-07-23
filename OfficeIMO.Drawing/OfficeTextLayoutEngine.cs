using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free text layout helpers for OfficeIMO renderers.
/// </summary>
public static partial class OfficeTextLayoutEngine {
    private const int DefaultTabSize = 4;
    private const int MaximumLayoutTextCharacters = 100_000;
    private const int MaximumLayoutLines = 4_096;
    private const int MaximumLayoutTextRuns = 4_096;
    private const int MaximumExhaustiveEllipsisElements = 256;

    /// <summary>
    /// Wraps text into measured lines using the supplied measurement delegate.
    /// </summary>
    /// <param name="text">Text to wrap. CRLF and CR line breaks are normalized to LF.</param>
    /// <param name="fontSize">Font size passed to <paramref name="measure"/>.</param>
    /// <param name="maxWidth">Maximum line width before wrapping or breaking long words.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while wrapping.</param>
    /// <returns>Measured wrapped lines. Empty input returns one empty line.</returns>
    public static IReadOnlyList<OfficeTextLine> WrapLines(string? text, double fontSize, double maxWidth, Func<string?, double, double> measure, OfficeTextParagraphIndent? paragraphIndent = null) =>
        WrapLines(text, fontSize, maxWidth, measure, paragraphIndent, out _);

    private static IReadOnlyList<OfficeTextLine> WrapLines(
        string? text,
        double fontSize,
        double maxWidth,
        Func<string?, double, double> measure,
        OfficeTextParagraphIndent? paragraphIndent,
        out bool clipped) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string bounded = LimitLayoutText(text ?? string.Empty, out bool inputTruncated);
        string expanded = ExpandTabs(bounded);
        string value = LimitLayoutText(expanded, out bool expandedTruncated);
        clipped = inputTruncated || expandedTruncated;
        double width = Math.Max(0D, maxWidth);
        OfficeTextParagraphIndent indent = paragraphIndent ?? OfficeTextParagraphIndent.Empty;
        string[] sourceLines = value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var output = new List<OfficeTextLine>();
        foreach (string sourceLine in sourceLines) {
            string line = sourceLine;
            bool firstVisualLine = true;
            if (line.Length == 0 || IsWhitespaceRun(line)) {
                if (output.Count >= MaximumLayoutLines) {
                    clipped = true;
                    return output;
                }

                output.Add(new OfficeTextLine(string.Empty, 0D, ResolveLineOffset(indent, firstVisualLine)));
                continue;
            }

            string current = string.Empty;
            foreach (PlainTextToken token in CreatePlainTextTokens(line)) {
                double currentOffset = ResolveLineOffset(indent, firstVisualLine);
                double currentWidth = Math.Max(0D, width - currentOffset);
                if (token.IsWhitespace) {
                    if (current.Length == 0) {
                        if (Measure(token.Text, fontSize, measure) <= currentWidth) {
                            current = token.Text;
                        }

                        continue;
                    }

                    string whitespaceCandidate = current + token.Text;
                    if (Measure(whitespaceCandidate, fontSize, measure) <= currentWidth) {
                        current = whitespaceCandidate;
                    }

                    continue;
                }

                if (Measure(token.Text, fontSize, measure) > currentWidth) {
                    string emitted = TrimTrailingSoftWrapWhitespace(current);
                    if (emitted.Length > 0) {
                        if (output.Count >= MaximumLayoutLines) {
                            clipped = true;
                            return output;
                        }

                        output.Add(CreateMeasuredLine(emitted, fontSize, measure, currentOffset));
                        firstVisualLine = false;
                    }

                    current = string.Empty;
                    foreach (OfficeTextLine part in BreakWord(token.Text, fontSize, width, measure, indent, firstVisualLine)) {
                        if (output.Count >= MaximumLayoutLines) {
                            clipped = true;
                            return output;
                        }

                        output.Add(part);
                        firstVisualLine = false;
                    }

                    continue;
                }

                string candidate = current + token.Text;
                if (current.Length > 0 && Measure(candidate, fontSize, measure) > currentWidth) {
                    string emitted = TrimTrailingSoftWrapWhitespace(current);
                    if (emitted.Length > 0) {
                        if (output.Count >= MaximumLayoutLines) {
                            clipped = true;
                            return output;
                        }

                        output.Add(CreateMeasuredLine(emitted, fontSize, measure, currentOffset));
                        firstVisualLine = false;
                    }

                    current = token.Text;
                } else {
                    current = candidate;
                }
            }

            if (current.Length > 0) {
                string emitted = TrimTrailingSoftWrapWhitespace(current);
                if (emitted.Length > 0) {
                    if (output.Count >= MaximumLayoutLines) {
                        clipped = true;
                        return output;
                    }

                    output.Add(CreateMeasuredLine(emitted, fontSize, measure, ResolveLineOffset(indent, firstVisualLine)));
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

    private static string LimitLayoutText(string text) {
        return LimitLayoutText(text, out _);
    }

    private static string LimitLayoutText(string text, out bool truncated) {
        truncated = text.Length > MaximumLayoutTextCharacters;
        if (!truncated) return text;
        int length = MaximumLayoutTextCharacters;
        if (length > 0 && char.IsHighSurrogate(text[length - 1])) length--;
        return text.Substring(0, length);
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
            max = Math.Max(max, lines[i].OffsetX + lines[i].Width);
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

        string bounded = LimitLayoutText(text ?? string.Empty, out bool inputTruncated);
        string value = LimitLayoutText(ExpandTabs(bounded), out bool expandedTruncated);
        bool limitTruncated = inputTruncated || expandedTruncated;
        double width = Math.Max(0D, maxWidth);
        double measured = Measure(value, fontSize, measure);
        if (measured <= width) {
            clipped = limitTruncated;
            return new OfficeTextLine(value, measured);
        }

        clipped = true;
        const string ellipsis = "...";
        value = FindFittingTextElementSlice(value, fontSize, width, measure, trimStart: false, ellipsis);
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

        string bounded = LimitLayoutText(text ?? string.Empty, out bool inputTruncated);
        string value = LimitLayoutText(ExpandTabs(bounded), out bool expandedTruncated);
        bool limitTruncated = inputTruncated || expandedTruncated;
        double width = Math.Max(0D, maxWidth);
        double measured = Measure(value, fontSize, measure);
        if (measured <= width) {
            clipped = limitTruncated;
            return new OfficeTextLine(value, measured);
        }

        clipped = true;
        const string ellipsis = "...";
        value = FindFittingTextElementSlice(value, fontSize, width, measure, trimStart: true, ellipsis);
        return new OfficeTextLine(value, Measure(value, fontSize, measure));
    }

    private static string FindFittingTextElementSlice(
        string value,
        double fontSize,
        double width,
        Func<string?, double, double> measure,
        bool trimStart,
        string ellipsis) {
        if (Measure(ellipsis, fontSize, measure) > width) return string.Empty;
        int[] starts = StringInfo.ParseCombiningCharacters(value);
        double MeasureCandidate(int elementCount) =>
            Measure(CreateEllipsizedSlice(value, starts, elementCount, trimStart, ellipsis), fontSize, measure);

        // Measurement delegates may be non-monotonic because of kerning or caller-defined
        // behavior. Preserve the legacy largest-fitting result for ordinary line lengths,
        // while keeping adversarial long-line trimming logarithmic.
        if (starts.Length <= MaximumExhaustiveEllipsisElements) {
            for (int elementCount = starts.Length; elementCount >= 0; elementCount--) {
                if (MeasureCandidate(elementCount) <= width) {
                    return CreateEllipsizedSlice(value, starts, elementCount, trimStart, ellipsis);
                }
            }

            return ellipsis;
        }

        int low = 0;
        int high = starts.Length;
        while (low < high) {
            int middle = low + ((high - low + 1) / 2);
            if (MeasureCandidate(middle) <= width) {
                low = middle;
            } else {
                high = middle - 1;
            }
        }

        const int monotonicityVerificationElements = 8;
        bool nonMonotonic = false;
        double previousWidth = MeasureCandidate(low);
        int verificationEnd = Math.Min(starts.Length, low + monotonicityVerificationElements);
        for (int elementCount = low + 1; elementCount <= verificationEnd; elementCount++) {
            double candidateWidth = MeasureCandidate(elementCount);
            if (candidateWidth < previousWidth) {
                nonMonotonic = true;
            }
            if (candidateWidth <= width) {
                low = elementCount;
            }
            previousWidth = candidateWidth;
        }

        if (nonMonotonic) {
            low = 0;
            for (int elementCount = 0; elementCount <= starts.Length; elementCount++) {
                if (MeasureCandidate(elementCount) <= width) {
                    low = elementCount;
                }
            }
        }

        return CreateEllipsizedSlice(value, starts, low, trimStart, ellipsis);
    }

    private static string CreateEllipsizedSlice(string value, int[] starts, int elementCount, bool trimStart, string ellipsis) {
        if (elementCount <= 0) return ellipsis;
        if (trimStart) {
            int start = elementCount >= starts.Length ? 0 : starts[starts.Length - elementCount];
            return ellipsis + value.Substring(start);
        }

        int length = elementCount >= starts.Length ? value.Length : starts[elementCount];
        return value.Substring(0, length) + ellipsis;
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
        string value = ExpandTabs(text ?? string.Empty);
        if (value.Length == 0 || Measure(value, resolvedFontSize, measure) <= width) {
            return resolvedFontSize;
        }

        if (Measure(value, minFontSize, measure) > width) {
            return minFontSize;
        }

        double low = minFontSize;
        double high = resolvedFontSize;
        int count = Math.Max(1, iterations);
        for (int i = 0; i < count; i++) {
            double candidate = (low + high) / 2D;
            if (Measure(value, candidate, measure) <= width) {
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
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped text.</param>
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
        bool shrinkToFit = false,
        OfficeTextParagraphIndent? paragraphIndent = null) =>
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
            OfficeTextOverflowBehavior.Ellipsis,
            paragraphIndent);

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
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped text.</param>
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
        OfficeTextOverflowBehavior overflowBehavior,
        OfficeTextParagraphIndent? paragraphIndent = null) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        string sourceText = text ?? string.Empty;
        string boundedText = LimitLayoutText(sourceText);
        string expandedText = ExpandTabs(forceSingleLine ? NormalizeSingleLineText(boundedText) : boundedText);
        string layoutText = LimitLayoutText(expandedText);
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
        bool clipped = boundedText.Length < sourceText.Length || layoutText.Length < expandedText.Length;

        OfficeTextParagraphIndent indent = paragraphIndent ?? OfficeTextParagraphIndent.Empty;
        if (effectiveWrap) {
            lines = WrapLines(layoutText, layoutFontSize, width, measure, indent, out bool lineLimitClipped);
            clipped |= lineLimitClipped;
        } else {
            string normalized = layoutText.Replace("\r\n", "\n").Replace('\r', '\n');
            int lineBreak = normalized.IndexOf('\n');
            string firstLine = lineBreak >= 0 ? normalized.Substring(0, lineBreak) : normalized;
            double offset = ResolveLineOffset(indent, firstVisualLine: true);
            OfficeTextLine line = ResolveOverflowLine(firstLine, layoutFontSize, Math.Max(0D, width - offset), measure, overflowBehavior, out bool lineClipped, offset);
            clipped |= lineClipped;
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
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while fitting wrapped text.</param>
    /// <returns>Measured text block with the resolved font size, line height, width, height, and lines.</returns>
    public static OfficeTextBlockLayout FitWrappedText(
        string? text,
        double fontSize,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        double minimumFontSize,
        Func<string?, double, double> measure,
        OfficeTextParagraphIndent? paragraphIndent = null) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedFontSize = NormalizePositive(fontSize, 1D);
        double minFontSize = Math.Max(1D, NormalizePositive(minimumFontSize, 1D));
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        OfficeTextParagraphIndent indent = paragraphIndent ?? OfficeTextParagraphIndent.Empty;
        OfficeTextBlockLayout layout = CreateBlockLayout(text, resolvedFontSize, width, lineFactor, measure, indent);
        double scaleDown = Math.Min(1D, Math.Min(width / Math.Max(layout.Width, 1D), height / Math.Max(layout.Height, 1D)));
        if (scaleDown < 0.98D) {
            resolvedFontSize = Math.Max(minFontSize, resolvedFontSize * Math.Max(0D, scaleDown));
            layout = CreateBlockLayout(text, resolvedFontSize, width, lineFactor, measure, indent);
        }

        return ClipTextBlockToHeight(
            layout.Lines,
            layout.FontSize,
            layout.LineHeight,
            width,
            height,
            measure,
            layout.Clipped);
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
                    visible[visible.Count - 1] = ApplyLineOffset(TrimLineToWidth(last.Text + "...", resolvedFontSize, Math.Max(0D, width - last.OffsetX), measure, out _), last.OffsetX);
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

    private static IEnumerable<OfficeTextLine> BreakWord(string word, double fontSize, double maxWidth, Func<string?, double, double> measure, OfficeTextParagraphIndent paragraphIndent, bool firstVisualLine) {
        int[] elementStarts = StringInfo.ParseCombiningCharacters(word);
        IReadOnlyList<int> preferredBreaks = OfficeTextLineBreaks.GetBreakPositions(word);
        int position = 0;
        int elementIndex = 0;
        while (position < word.Length && elementIndex < elementStarts.Length) {
            double offset = ResolveLineOffset(paragraphIndent, firstVisualLine);
            double width = Math.Max(0D, maxWidth - offset);
            int fittingCount = FindFittingElementCount(
                word,
                elementStarts,
                elementIndex,
                position,
                fontSize,
                width,
                measure);
            int fittingBoundary = GetTextElementEnd(elementStarts, word.Length, elementIndex + fittingCount - 1);
            int fittingPreferred = FindLastBreakPosition(preferredBreaks, position, fittingBoundary);
            int selected = fittingPreferred > position ? fittingPreferred : fittingBoundary;
            if (selected <= position) {
                break;
            }

            string part = word.Substring(position, selected - position);
            yield return new OfficeTextLine(part, Measure(part, fontSize, measure), offset);
            position = selected;
            if (selected >= word.Length) {
                elementIndex = elementStarts.Length;
            } else {
                int nextIndex = Array.BinarySearch(elementStarts, selected);
                elementIndex = nextIndex >= 0 ? nextIndex : ~nextIndex;
            }
            firstVisualLine = false;
        }
    }

    private static int FindFittingElementCount(
        string word,
        int[] elementStarts,
        int elementIndex,
        int position,
        double fontSize,
        double width,
        Func<string?, double, double> measure) {
        int remaining = elementStarts.Length - elementIndex;
        int lower = 0;
        int upper = 1;
        while (upper <= remaining) {
            int boundary = GetTextElementEnd(elementStarts, word.Length, elementIndex + upper - 1);
            if (Measure(word.Substring(position, boundary - position), fontSize, measure) > width) break;
            lower = upper;
            if (upper == remaining) return upper;
            upper = Math.Min(remaining, checked(upper * 2));
        }

        if (lower == 0 && upper == 1) return 1;
        int high = Math.Min(remaining, upper);
        while (lower + 1 < high) {
            int middle = lower + ((high - lower) / 2);
            int boundary = GetTextElementEnd(elementStarts, word.Length, elementIndex + middle - 1);
            if (Measure(word.Substring(position, boundary - position), fontSize, measure) <= width) {
                lower = middle;
            } else {
                high = middle;
            }
        }

        return Math.Max(1, lower);
    }

    private static int GetTextElementEnd(int[] starts, int textLength, int elementIndex) =>
        elementIndex + 1 < starts.Length ? starts[elementIndex + 1] : textLength;

    private static int FindLastBreakPosition(IReadOnlyList<int> positions, int exclusiveMinimum, int inclusiveMaximum) {
        int result = -1;
        int low = 0;
        int high = positions.Count - 1;
        while (low <= high) {
            int middle = low + ((high - low) / 2);
            int value = positions[middle];
            if (value <= inclusiveMaximum) {
                if (value > exclusiveMinimum) result = value;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }

        return result;
    }

    private static OfficeTextLine ResolveOverflowLine(
        string? text,
        double fontSize,
        double maxWidth,
        Func<string?, double, double> measure,
        OfficeTextOverflowBehavior overflowBehavior,
        out bool clipped,
        double offsetX = 0D) {
        string value = ExpandTabs(text ?? string.Empty);
        double width = Math.Max(0D, maxWidth);
        double measured = Measure(value, fontSize, measure);
        if (measured <= width || overflowBehavior != OfficeTextOverflowBehavior.Clip) {
            return ApplyLineOffset(TrimLineToWidth(value, fontSize, width, measure, out clipped), offsetX);
        }

        clipped = true;
        return new OfficeTextLine(value, measured, offsetX);
    }

    private static double Measure(string? text, double fontSize, Func<string?, double, double> measure) =>
        string.IsNullOrEmpty(text) ? 0D : measure(text, fontSize);

    private static string NormalizeSingleLineText(string text) =>
        text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');

    private static string ExpandTabs(string text, int tabSize = DefaultTabSize) {
        if (string.IsNullOrEmpty(text) || text.IndexOf('\t') < 0) {
            return text;
        }

        int resolvedTabSize = Math.Max(1, tabSize);
        var builder = new StringBuilder(text.Length);
        int column = 0;
        for (int i = 0; i < text.Length; i++) {
            char value = text[i];
            if (value == '\t') {
                int spaces = resolvedTabSize - (column % resolvedTabSize);
                builder.Append(' ', spaces);
                column += spaces;
                continue;
            }

            builder.Append(value);
            if (value == '\r' || value == '\n') {
                column = 0;
            } else {
                column++;
            }
        }

        return builder.ToString();
    }

    private static IEnumerable<PlainTextToken> CreatePlainTextTokens(string text) {
        var token = new StringBuilder();
        bool? tokenWhitespace = null;
        for (int i = 0; i < text.Length; i++) {
            char value = text[i];
            bool isWhitespace = char.IsWhiteSpace(value);
            if (tokenWhitespace.HasValue && tokenWhitespace.Value != isWhitespace) {
                yield return new PlainTextToken(token.ToString(), tokenWhitespace.Value);
                token.Clear();
            }

            token.Append(value);
            tokenWhitespace = isWhitespace;
        }

        if (token.Length > 0) {
            yield return new PlainTextToken(token.ToString(), tokenWhitespace.GetValueOrDefault());
        }
    }

    private static OfficeTextLine CreateMeasuredLine(string text, double fontSize, Func<string?, double, double> measure, double offsetX = 0D) =>
        new OfficeTextLine(text, Measure(text, fontSize, measure), offsetX);

    private static OfficeTextLine ApplyLineOffset(OfficeTextLine line, double offsetX) =>
        new OfficeTextLine(line.Text, line.Width, offsetX);

    private static double ResolveLineOffset(OfficeTextParagraphIndent paragraphIndent, bool firstVisualLine) =>
        firstVisualLine ? paragraphIndent.FirstLineOffset : paragraphIndent.ContinuationLineOffset;

    private static OfficeTextBlockLayout CreateBlockLayout(string? text, double fontSize, double maxWidth, double lineHeightFactor, Func<string?, double, double> measure, OfficeTextParagraphIndent paragraphIndent) {
        IReadOnlyList<OfficeTextLine> lines = WrapLines(text, fontSize, maxWidth, measure, paragraphIndent, out bool clipped);
        double lineHeight = fontSize * lineHeightFactor;
        double width = MeasureMaxLineWidth(lines);
        double height = Math.Max(fontSize, ((lines.Count - 1) * lineHeight) + fontSize);
        return new OfficeTextBlockLayout(lines, fontSize, lineHeight, width, height, clipped);
    }

    private static double NormalizePositive(double value, double fallback) =>
        value > 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && !double.IsNaN(value) ? value : 0D;

    private readonly struct PlainTextToken {
        internal PlainTextToken(string text, bool isWhitespace) {
            Text = text;
            IsWhitespace = isWhitespace;
        }

        internal string Text { get; }

        internal bool IsWhitespace { get; }
    }
}
