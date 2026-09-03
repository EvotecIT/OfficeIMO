using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight layout utilities to group text spans into lines and infer multi-column reading order.
/// Zero-dependency and heuristic by design.
/// </summary>
internal static class TextLayoutEngine {
    public sealed class Options {
        /// <summary>Assume page margins (points) when inferring columns. Default: 36 pt (0.5").</summary>
        public double MarginLeft { get; set; } = 36;
        public double MarginRight { get; set; } = 36;
        /// <summary>Histogram bin width for gutter detection. Default: 5 pt.</summary>
        public double BinWidth { get; set; } = 5;
        /// <summary>Minimum gutter width to consider split into two columns. Default: 24 pt.</summary>
        public double MinGutterWidth { get; set; } = 24;
        /// <summary>Maximum Y delta (as fraction of font size) to group spans into the same line. Default: 0.6.</summary>
        public double LineMergeToleranceEm { get; set; } = 0.6;
        /// <summary>Maximum absolute Y delta (points) to merge spans into the same line. Default: 2.5.</summary>
        public double LineMergeMaxPoints { get; set; } = 2.5;
        /// <summary>Force single column when true; skip gutter detection.</summary>
        public bool ForceSingleColumn { get; set; }
        /// <summary>Horizontal direction for line construction and column emission.</summary>
        public PdfReadingDirection ReadingDirection { get; set; }
        /// <summary>Threshold in em units to insert a space between adjacent spans on the same line. Default: 0.35.</summary>
        public double GapSpaceThresholdEm { get; set; } = 0.35;
        /// <summary>Threshold as a fraction of previous span's average glyph advance to insert a space. Default: 0.60.</summary>
        public double GapGlyphFactor { get; set; } = 0.60;
        /// <summary>When true, same-baseline spans separated by a wide gutter are emitted as separate lines.</summary>
        internal bool SplitWideSameBaselineRuns { get; set; }
    }

    public sealed class TextLine {
        public double Y { get; }
        public double XStart { get; }
        public double XEnd { get; }
        public string Text { get; }
        public IReadOnlyList<PdfTextSpan> Spans { get; }
        public int LogicalLineBreaksBefore { get; }
        public TextLine(double y, double xs, double xe, string text, List<PdfTextSpan> spans) {
            Y = y; XStart = xs; XEnd = xe; Text = text; Spans = spans;
            LogicalLineBreaksBefore = spans.Count == 0 ? 0 : spans.Max(span => span.LogicalLineBreaksBefore);
        }
    }

    public sealed class ColumnLayout {
        public (double From, double To) Left { get; }
        public (double From, double To) Right { get; }
        public bool IsTwoColumns { get; }
        public ColumnLayout((double, double) left, (double, double) right, bool two) { Left = left; Right = right; IsTwoColumns = two; }
    }

    /// <summary>
    /// Split lines into horizontal bands (blocks) based on Y gaps.
    /// Useful for de-duplicating and for column/table detection within local neighborhoods.
    /// </summary>
    public static List<List<TextLine>> BandLines(List<TextLine> lines, Options? options = null) =>
        BandLines(lines, options, consumeWork: null, cancellationCheck: null);

    internal static List<List<TextLine>> BandLines(
        List<TextLine> lines,
        Options? options,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        options ??= new Options();
        var result = new List<List<TextLine>>();
        if (lines.Count == 0) return result;
        cancellationCheck?.Invoke();
        consumeWork?.Invoke(lines.Count);
        // Work on lines sorted by Y desc
        var ordered = lines.OrderByDescending(l => l.Y).ToList();
        cancellationCheck?.Invoke();
        // Band gap: larger than intra-line tolerance to group adjacent lines sensibly
        double baseGap = Math.Max(8.0, options.LineMergeMaxPoints * 3.0);
        var current = new List<TextLine>();
        double currentY = ordered[0].Y;
        foreach (var ln in ordered) {
            cancellationCheck?.Invoke();
            if (current.Count == 0) { current.Add(ln); currentY = ln.Y; continue; }
            if (Math.Abs(ln.Y - currentY) <= baseGap) {
                current.Add(ln);
            } else {
                result.Add(current);
                current = new List<TextLine> { ln };
                currentY = ln.Y;
            }
        }
        if (current.Count > 0) result.Add(current);
        return result;
    }

    /// <summary>Builds text lines from spans using Y-clustering and X-sorting.</summary>
    public static List<TextLine> BuildLines(IReadOnlyList<PdfTextSpan> spans, Options? options = null) =>
        BuildLines(spans, options, consumeWork: null, cancellationCheck: null);

    internal static List<TextLine> BuildLines(
        IReadOnlyList<PdfTextSpan> spans,
        Options? options,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        options ??= new Options();
        if (spans.Count == 0) return new List<TextLine>();
        cancellationCheck?.Invoke();
        consumeWork?.Invoke(spans.Count);
        // Sort by Y desc, then X asc
        var ordered = spans.OrderByDescending(s => s.Y).ThenBy(s => s.X).ToList();
        cancellationCheck?.Invoke();
        // Estimate avg font size (robust median)
        double medianSize = Median(ordered.Select(s => s.FontSize));
        var lines = new List<TextLine>();
        var current = new List<PdfTextSpan>();
        double currentY = ordered[0].Y;
        double currentFont = ordered[0].FontSize;
        foreach (var s in ordered) {
            cancellationCheck?.Invoke();
            if (current.Count == 0) { current.Add(s); currentY = s.Y; continue; }
            double tolAbs = Math.Min(options.LineMergeMaxPoints, Math.Min(currentFont, s.FontSize) * options.LineMergeToleranceEm);
            if (tolAbs < 0.5) tolAbs = 0.5;
            if (Math.Abs(s.Y - currentY) <= tolAbs) {
                current.Add(s);
                currentFont = (currentFont * (current.Count - 1) + s.FontSize) / current.Count;
            } else {
                AddBuiltLines(lines, current, options);
                current.Clear();
                current.Add(s);
                currentY = s.Y; currentFont = s.FontSize;
            }
        }
        if (current.Count > 0) AddBuiltLines(lines, current, options);
        // Drop obvious duplicate lines drawn twice at the same Y (e.g., shadow/overprint)
        lines = DeduplicateLines(lines, consumeWork, cancellationCheck);
        PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
            options.ReadingDirection,
            spans.OrderBy(static span => span.ContentOrderKey)
                .ThenBy(static span => span.PaintOrder)
                .Select(static span => span.Text));
        lines.Sort((left, right) => {
            int baseline = right.Y.CompareTo(left.Y);
            return baseline != 0
                ? baseline
                : direction == PdfReadingDirection.RightToLeft
                    ? right.XStart.CompareTo(left.XStart)
                    : left.XStart.CompareTo(right.XStart);
        });
        cancellationCheck?.Invoke();
        return lines;
    }

    internal static IReadOnlyList<PdfTextSpan> FilterIgnoredPageBands(
        IReadOnlyList<PdfTextSpan> spans,
        PdfReadPage page,
        PdfTextLayoutOptions options) =>
        FilterIgnoredPageBands(spans, page, options, consumeWork: null, cancellationCheck: null);

    internal static IReadOnlyList<PdfTextSpan> FilterIgnoredPageBands(
        IReadOnlyList<PdfTextSpan> spans,
        PdfReadPage page,
        PdfTextLayoutOptions options,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (options.IgnoreHeaderHeight <= 0D && options.IgnoreFooterHeight <= 0D) return spans;
        double visualHeight = page.GetVisualPageSize().Height;
        Matrix2D visualTransform = page.GetVisualPageTransform();
        double bodyTop = options.IgnoreHeaderHeight;
        double bodyBottom = visualHeight - options.IgnoreFooterHeight;
        var filtered = new List<PdfTextSpan>(spans.Count);
        for (int index = 0; index < spans.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            PdfTextSpan span = spans[index];
            double visualBaselineFromBottom = visualTransform.Transform(span.X, span.Y).Y;
            double visualBaselineFromTop = visualHeight - visualBaselineFromBottom;
            if ((options.IgnoreHeaderHeight <= 0D || visualBaselineFromTop > bodyTop) &&
                (options.IgnoreFooterHeight <= 0D || visualBaselineFromTop < bodyBottom)) {
                filtered.Add(span);
            }
        }
        return Array.AsReadOnly(filtered.ToArray());
    }

    /// <summary>Attempts to detect a two-column layout by finding a vertical low-coverage gutter.</summary>
    public static ColumnLayout DetectColumns(List<TextLine> lines, double pageWidth, Options? options = null) {
        options ??= new Options();
        if (options.ForceSingleColumn || lines.Count == 0 || pageWidth <= 0) {
            return new ColumnLayout((options.MarginLeft, pageWidth - options.MarginRight), (0, 0), false);
        }
        int bins = (int)Math.Max(1, Math.Ceiling(pageWidth / options.BinWidth));
        var hist = new int[bins];
        void AddCoverage(double xs, double xe) {
            int b0 = Clamp((int)Math.Floor(xs / options.BinWidth), 0, bins - 1);
            int b1 = Clamp((int)Math.Floor(xe / options.BinWidth), 0, bins - 1);
            for (int b = b0; b <= b1; b++) hist[b]++;
        }
        foreach (var ln in lines) {
            AddCoverage(ln.XStart, ln.XEnd);
        }
        // Identify longest low-coverage run near middle of page
        int mid = bins / 2;
        int bestStart = -1, bestEnd = -1, bestLen = 0;
        int curStart = -1;
        int maxVal = hist.Length == 0 ? 0 : hist.Max();
        // threshold: bins with less than 10% of max coverage
        double thr = maxVal * 0.1;
        for (int i = 0; i < bins; i++) {
            bool low = hist[i] <= thr;
            if (low) {
                if (curStart < 0) curStart = i;
            } else if (curStart >= 0) {
                int curEnd = i - 1;
                int curLen = curEnd - curStart + 1;
                if (curLen > bestLen && Math.Abs(((curStart + curEnd) / 2) - mid) < bins * 0.25) {
                    bestLen = curLen; bestStart = curStart; bestEnd = curEnd;
                }
                curStart = -1;
            }
        }
        if (curStart >= 0) {
            int curEnd = bins - 1;
            int curLen = curEnd - curStart + 1;
            if (curLen > bestLen && Math.Abs(((curStart + curEnd) / 2) - mid) < bins * 0.25) {
                bestLen = curLen; bestStart = curStart; bestEnd = curEnd;
            }
        }
        if (bestLen * options.BinWidth >= options.MinGutterWidth) {
            double gutterL = bestStart * options.BinWidth;
            double gutterR = (bestEnd + 1) * options.BinWidth;
            var left = (options.MarginLeft, Math.Max(options.MarginLeft, gutterL));
            var right = (Math.Min(pageWidth - options.MarginRight, gutterR), pageWidth - options.MarginRight);
            return new ColumnLayout(left, right, true);
        }
        return new ColumnLayout((options.MarginLeft, pageWidth - options.MarginRight), (0, 0), false);
    }

    /// <summary>
    /// Emits text in inferred reading order. For two columns: left column top→bottom, then right.
    /// For single column: top→bottom.
    /// </summary>
    public static string EmitText(List<TextLine> lines, ColumnLayout columns, PdfTextLayoutOptions? options = null) {
        var sb = new StringBuilder();
        PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
            options?.ReadingDirection ?? PdfReadingDirection.Auto,
            lines.OrderByDescending(static line => line.Y).Select(static line => line.Text));
        if (columns.IsTwoColumns) {
            var left = lines.Where(l => l.XStart >= columns.Left.From && l.XStart <= columns.Left.To).OrderByDescending(l => l.Y);
            var right = lines.Where(l => l.XStart >= columns.Right.From && l.XStart <= columns.Right.To).OrderByDescending(l => l.Y);
            bool first = true;
            IEnumerable<IEnumerable<TextLine>> columnsInReadingOrder = direction == PdfReadingDirection.RightToLeft
                ? new[] { right, left }
                : new[] { left, right };
            foreach (IEnumerable<TextLine> column in columnsInReadingOrder) {
                foreach (TextLine line in column) AppendLine(sb, line, ref first);
            }
        } else {
            bool first = true;
            foreach (var ln in lines.OrderByDescending(l => l.Y)) {
                AppendLine(sb, ln, ref first);
            }
        }
        string text = sb.ToString();
        if (options?.JoinSoftHyphensAcrossLines == true) {
            text = JoinHyphenation(text);
        }
        return text;
    }

    private static void AppendLine(StringBuilder builder, TextLine line, ref bool first) {
        if (!first) {
            int lineBreaks = Math.Max(1, line.LogicalLineBreaksBefore);
            builder.Append('\n', lineBreaks);
        }

        builder.Append(line.Text);
        first = false;
    }

    private static string JoinHyphenation(string text) {
        // A soft hyphen is an explicit discretionary-break signal. A visible hyphen is authored
        // content and cannot be removed safely from PDF geometry or letter case alone.
        return System.Text.RegularExpressions.Regex.Replace(text, "\u00AD\n", string.Empty);
    }

    private static void AddBuiltLines(List<TextLine> lines, List<PdfTextSpan> spans, Options options) {
        if (!options.SplitWideSameBaselineRuns || spans.Count <= 1) {
            lines.Add(BuildLine(spans, options));
            return;
        }

        foreach (var run in SplitWideSameBaselineRuns(spans, options)) {
            lines.Add(BuildLine(run, options));
        }
    }

    private static List<List<PdfTextSpan>> SplitWideSameBaselineRuns(List<PdfTextSpan> spans, Options options) {
        var ordered = spans.OrderBy(s => s.X).ToList();
        var runs = new List<List<PdfTextSpan>>();
        var current = new List<PdfTextSpan> { ordered[0] };
        double minimumRunGap = Math.Max(12, options.MinGutterWidth);

        for (int i = 1; i < ordered.Count; i++) {
            var previous = ordered[i - 1];
            var span = ordered[i];
            double previousEnd = previous.X + Math.Max(0, previous.Advance);
            double gap = span.X - previousEnd;
            if (gap >= minimumRunGap) {
                runs.Add(current);
                current = new List<PdfTextSpan>();
            }

            current.Add(span);
        }

        if (current.Count > 0) {
            runs.Add(current);
        }

        return runs;
    }

    private static TextLine BuildLine(List<PdfTextSpan> spans, Options? options) {
        List<PdfTextSpan> sourceOrder = spans
            .OrderBy(static span => span.ContentOrderKey)
            .ThenBy(static span => span.PaintOrder)
            .ToList();
        PdfReadingDirection direction = PdfTextDirectionAnalysis.Resolve(
            options?.ReadingDirection ?? PdfReadingDirection.Auto,
            sourceOrder.Select(static span => span.Text));
        // Keep stored geometry left-to-right for table/cell detection while emitting line text
        // in the resolved writing direction.
        spans.Sort(static (left, right) => left.X.CompareTo(right.X));
        List<PdfTextSpan> textSpans = direction == PdfReadingDirection.RightToLeft
            ? spans.AsEnumerable().Reverse().ToList()
            : spans;
        bool hasExplicitWhitespace = spans.Any(span =>
            ContainsWhitespace(span.Text) ||
            span.LogicalLeadingSpace ||
            span.LogicalTrailingSpace);
        double xs = spans.Min(static span => span.X);
        double xe = spans.Max(static span => span.X + Math.Max(0D, span.Advance));
        var text = new StringBuilder();
        for (int i = 0; i < textSpans.Count; i++) {
            var s = textSpans[i];
            if (i > 0) {
                var previous = textSpans[i - 1];
                bool explicitBoundarySpace = previous.LogicalTrailingSpace || s.LogicalLeadingSpace;
                if (explicitBoundarySpace && text.Length > 0 && text[text.Length - 1] != ' ') {
                    text.Append(' ');
                }

                // Add a space heuristically if large X gap between spans
                var prev = previous;
                double gap = direction == PdfReadingDirection.RightToLeft
                    ? prev.X - (s.X + Math.Max(0D, s.Advance))
                    : s.X - (prev.X + Math.Max(0D, prev.Advance));
                // dynamic threshold based on previous span's average glyph advance
                double prevAvg = SafeAvgAdvance(prev);
                double glyphFactor = options?.GapGlyphFactor ?? 0.6;
                double glyphThreshold = prevAvg * glyphFactor;
                // fallback to em threshold when prevAvg unavailable
                double emThreshold = (options?.GapSpaceThresholdEm ?? 0.25) * s.FontSize;
                double threshold = Math.Max(emThreshold, glyphThreshold);
                bool isLeader = IsLeaderRun(s.Text);
                // Tight word-join rule: letters adjacent use stricter threshold (slightly more permissive)
                if (!explicitBoundarySpace && IsWordJoin(prev.Text, s.Text)) {
                    // be less aggressive: add space whenever gap exceeds ~0.65x glyph-advance or 0.30em
                    double tight = System.Math.Max(1.0, System.Math.Min(3.0, System.Math.Min(prevAvg * 0.65, s.FontSize * 0.30)));
                    if (gap > tight) text.Append(' ');
                    else {
                        // Fallback: if both look like full words and there is a visible gap, insert a space
                        bool bothAlphaLong = AllWordish(prev.Text) && AllWordish(s.Text) &&
                            PdfUnicodeScalarAnalysis.CountScalars(prev.Text) >= 2 &&
                            PdfUnicodeScalarAnalysis.CountScalars(s.Text) >= 2;
                        if ((bothAlphaLong || ShouldRespectVisibleGap(prev.Text, s.Text)) && IsVisibleWordGap(gap, s.FontSize) && (text.Length > 0 && text[text.Length - 1] != ' ')) text.Append(' ');
                    }
                } else if (!explicitBoundarySpace && !isLeader) {
                    // Guard: if both chunks look like full words (>=2 letters) and there is any visible gap, emit a space
                    bool bothAlphaLong = AllWordish(prev.Text) && AllWordish(s.Text) &&
                        PdfUnicodeScalarAnalysis.CountScalars(prev.Text) >= 2 &&
                        PdfUnicodeScalarAnalysis.CountScalars(s.Text) >= 2;
                    if ((bothAlphaLong || ShouldRespectVisibleGap(prev.Text, s.Text)) && IsVisibleWordGap(gap, s.FontSize) && (text.Length > 0 && text[text.Length - 1] != ' ')) {
                        text.Append(' ');
                    }
                    if (gap > threshold) text.Append(' ');
                } else if (!explicitBoundarySpace) {
                    if (gap > 0 && text.Length > 0 && text[text.Length - 1] != ' ') text.Append(' '); // one space before leader
                }
            }
            // drop duplicate shadows: if same text repeats with almost no gap
            if (text.Length > 0 && IsSameAsTail(text, s.Text) && i > 0) {
                var prev = textSpans[i - 1];
                if (IsSubstantiallyOverlapping(prev, s)) {
                    continue;
                }
            }
            text.Append(s.Text);
            // if leader followed by number, ensure a single space
            if (IsLeaderRun(s.Text) && i + 1 < textSpans.Count && ContainsDigit(textSpans[i + 1].Text)) {
                if (text.Length > 0 && text[text.Length - 1] != ' ') text.Append(' ');
            }
        }
        string outText = text.ToString();
        if (!IsLeaderRun(outText)) {
            outText = hasExplicitWhitespace
                ? System.Text.RegularExpressions.Regex.Replace(outText, "\\s+", " ").Trim()
                : NormalizeLineText(outText);
        }
        return new TextLine(spans[0].Y, xs, xe, outText, new List<PdfTextSpan>(spans));
    }

    private static bool ContainsWhitespace(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsWhiteSpace(value[index])) {
                return true;
            }
        }

        return false;
    }

    private static double Median(IEnumerable<double> seq) {
        var list = seq.Where(v => v > 0).OrderBy(v => v).ToList();
        if (list.Count == 0) return 12;
        int mid = list.Count / 2;
        if (list.Count % 2 == 1) return list[mid];
        return (list[mid - 1] + list[mid]) / 2.0;
    }

    private static int Clamp(int v, int min, int max) => v < min ? min : (v > max ? max : v);

    private static List<TextLine> DeduplicateLines(
        List<TextLine> lines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (lines.Count <= 1) return lines;
        var result = new List<TextLine>(lines.Count);
        var used = new bool[lines.Count];
        for (int i = 0; i < lines.Count; i++) {
            cancellationCheck?.Invoke();
            if (used[i]) continue;
            var a = lines[i];
            result.Add(a);
            for (int j = i + 1; j < lines.Count; j++) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(1);
                if (used[j]) continue;
                var b = lines[j];
                // Near-identical baseline
                if (Math.Abs(a.Y - b.Y) <= 0.75) {
                    // Exact text match and significant X overlap => drop b
                    if (string.Equals(a.Text, b.Text, StringComparison.Ordinal)) {
                        double overlap = Math.Min(a.XEnd, b.XEnd) - Math.Max(a.XStart, b.XStart);
                        double len = Math.Max(1.0, Math.Min(a.XEnd - a.XStart, b.XEnd - b.XStart));
                        if (overlap / len > 0.6) { used[j] = true; continue; }
                        if (Math.Abs(a.XStart - b.XStart) <= 1.0) { used[j] = true; continue; }
                    }
                }
            }
        }
        return result;
    }

    private static bool IsLeaderRun(string s) {
        if (string.IsNullOrEmpty(s) || s.Length < 3) return false;
        char c = s[0];
        if (c != '.' && c != '-' && c != '_') return false;
        for (int i = 1; i < s.Length; i++) if (s[i] != c) return false; return true;
    }
    private static bool IsWordJoin(string left, string right) {
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)) return false;
        int a = GetLastScalar(left);
        int b = char.ConvertToUtf32(right, 0);
        bool aWord = PdfUnicodeScalarAnalysis.IsLastLetterOrDigit(left) || a is 0x29 or 0x22 or 0x27 or 0x2019;
        bool bWord = PdfUnicodeScalarAnalysis.IsFirstLetterOrDigit(right) || b is 0x28 or 0x22 or 0x27 or 0x2018;
        return aWord && bWord;
    }
    private static bool IsSameAsTail(StringBuilder sb, string s) {
        if (string.IsNullOrEmpty(s)) return false; int len = s.Length; if (sb.Length < len) return false;
        for (int i = 0; i < len; i++) if (sb[sb.Length - len + i] != s[i]) return false; return true;
    }
    private static bool IsSubstantiallyOverlapping(PdfTextSpan previous, PdfTextSpan current) {
        double previousWidth = Math.Max(0, previous.Advance);
        double currentWidth = Math.Max(0, current.Advance);
        if (previousWidth <= 0 || currentWidth <= 0) {
            return Math.Abs(previous.X - current.X) <= 0.8;
        }

        double previousEnd = previous.X + previousWidth;
        double currentEnd = current.X + currentWidth;
        double overlap = Math.Min(previousEnd, currentEnd) - Math.Max(previous.X, current.X);
        double narrowerWidth = Math.Min(previousWidth, currentWidth);
        return overlap > 0 && overlap / narrowerWidth >= 0.6;
    }
    private static bool ContainsDigit(string s) {
        return !string.IsNullOrEmpty(s) && PdfUnicodeScalarAnalysis.ContainsDecimalDigit(s);
    }
    private static bool AllWordish(string s) => PdfUnicodeScalarAnalysis.IsAllWordish(s);
    private static bool IsVisibleWordGap(double gap, double fontSize) =>
        gap > System.Math.Max(0.8, System.Math.Min(2.0, fontSize * 0.18));
    private static bool ShouldRespectVisibleGap(string left, string right) {
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)) return false;
        int a = GetLastScalar(left);
        int b = char.ConvertToUtf32(right, 0);
        bool leftBoundary = PdfUnicodeScalarAnalysis.IsLastLetterOrDigit(left) || a is 0x3A or 0x3B or 0x2C or 0x2E or 0x29 or 0x22 or 0x27 or 0x2019;
        bool rightBoundary = PdfUnicodeScalarAnalysis.IsFirstLetterOrDigit(right) || b is 0x28 or 0x22 or 0x27 or 0x2018;
        return leftBoundary && rightBoundary;
    }
    private static int GetLastScalar(string value) {
        int index = value.Length - 1;
        if (index > 0 && char.IsLowSurrogate(value[index]) && char.IsHighSurrogate(value[index - 1])) index--;
        return char.ConvertToUtf32(value, index);
    }
    private static string NormalizeLineText(string s) {
        if (string.IsNullOrEmpty(s)) return s;
        // Text content cannot prove that whitespace between separately painted glyph runs is an
        // intra-word fracture. Keep word boundaries here; geometric grouping owns gap decisions.
        return System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
    }
    private static double SafeAvgAdvance(PdfTextSpan span) {
        if (span.Advance <= 0) return span.FontSize * 0.5;
        int len = PdfUnicodeScalarAnalysis.CountScalars(span.Text ?? string.Empty); if (len <= 0) return span.FontSize * 0.5;
        return span.Advance / len;
    }
}

/// <summary>
/// Convenience helpers for callers to get column-aware text from a PdfReadPage.
/// </summary>
public static class PdfReadPageExtensions {
    /// <summary>
    /// Extracts text from a page with simple two-column detection when present.
    /// </summary>
    /// <param name="page">Source page.</param>
    /// <param name="options">Optional layout options controlling column detection, margins and trimming.</param>
    /// <returns>Plain text for this page in inferred reading order.</returns>
    public static string ExtractTextWithColumns(this PdfReadPage page, PdfTextLayoutOptions? options = null) {
        var spans = page.GetTextSpans();
        if (options is not null) {
            spans = TextLayoutEngine.FilterIgnoredPageBands(spans, page, options);
        }
        var engineOpts = options?.ToEngineOptions() ?? new TextLayoutEngine.Options();
        if (!engineOpts.ForceSingleColumn) {
            engineOpts.SplitWideSameBaselineRuns = true;
        }

        var lines = TextLayoutEngine.BuildLines(spans, engineOpts);
        var (w, _) = page.GetPageSize();
        var layout = TextLayoutEngine.DetectColumns(lines, w, engineOpts);
        return TextLayoutEngine.EmitText(lines, layout, options);
    }

    /// <summary>
    /// Extracts a simple structured model (lines, TOC entries, list items) for this page.
    /// </summary>
    /// <param name="page">Source page.</param>
    /// <param name="options">Optional layout options.</param>
    public static StructuredPage ExtractStructured(this PdfReadPage page, PdfTextLayoutOptions? options = null) {
        var spans = page.GetTextSpans();
        return ExtractStructured(page, spans, options, CancellationToken.None);
    }

    internal static StructuredPage ExtractStructured(
        this PdfReadPage page,
        IReadOnlyList<PdfTextSpan> spans,
        PdfTextLayoutOptions? options,
        CancellationToken cancellationToken,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null,
        IReadOnlyList<StructuredTable>? precomputedTables = null) {
        cancellationToken.ThrowIfCancellationRequested();
        cancellationCheck?.Invoke();
        var (_, pageHeight) = page.GetPageSize();
        if (options is not null) {
            spans = TextLayoutEngine.FilterIgnoredPageBands(
                spans,
                page,
                options,
                consumeWork,
                cancellationCheck ?? cancellationToken.ThrowIfCancellationRequested);
        }

        var engineOpts = options?.ToEngineOptions();
        StructuredPage result = ContentStructureExtractor.Extract(
            spans,
            engineOpts ?? new TextLayoutEngine.Options(),
            pageHeight,
            consumeWork,
            cancellationCheck ?? cancellationToken.ThrowIfCancellationRequested,
            precomputedTables);
        cancellationToken.ThrowIfCancellationRequested();
        cancellationCheck?.Invoke();
        return result;
    }
}
