using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight layout utilities to group text spans into lines and infer multi-column reading order.
/// Zero-dependency and heuristic by design.
/// </summary>
internal static class TextLayoutEngine {
    private static readonly HashSet<string> CommonSuffixes = new(StringComparer.OrdinalIgnoreCase) {
        "ion", "ions", "ing", "ment", "tion", "sion", "iation", "ization",
        "ability", "ality", "able", "ible", "ance", "ence", "al", "ally",
        "er", "ers", "ed", "ly", "ology", "ologies"
    };

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
        /// <summary>Threshold in em units to insert a space between adjacent spans on the same line. Default: 0.35.</summary>
        public double GapSpaceThresholdEm { get; set; } = 0.35;
        /// <summary>Threshold as a fraction of previous span's average glyph advance to insert a space. Default: 0.60.</summary>
        public double GapGlyphFactor { get; set; } = 0.60;
    }

    public sealed class TextLine {
        public double Y { get; }
        public double XStart { get; }
        public double XEnd { get; }
        public string Text { get; }
        public IReadOnlyList<PdfTextSpan> Spans { get; }
        public TextLine(double y, double xs, double xe, string text, List<PdfTextSpan> spans) {
            Y = y; XStart = xs; XEnd = xe; Text = text; Spans = spans;
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
    public static List<List<TextLine>> BandLines(List<TextLine> lines, Options? options = null) {
        options ??= new Options();
        var result = new List<List<TextLine>>();
        if (lines.Count == 0) return result;
        // Work on lines sorted by Y desc
        var ordered = lines.OrderByDescending(l => l.Y).ToList();
        // Band gap: larger than intra-line tolerance to group adjacent lines sensibly
        double baseGap = Math.Max(8.0, options.LineMergeMaxPoints * 3.0);
        var current = new List<TextLine>();
        double currentY = ordered[0].Y;
        foreach (var ln in ordered) {
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
    public static List<TextLine> BuildLines(IReadOnlyList<PdfTextSpan> spans, Options? options = null) {
        options ??= new Options();
        if (spans.Count == 0) return new List<TextLine>();
        // Sort by Y desc, then X asc
        var ordered = spans.OrderByDescending(s => s.Y).ThenBy(s => s.X).ToList();
        // Estimate avg font size (robust median)
        double medianSize = Median(ordered.Select(s => s.FontSize));
        var lines = new List<TextLine>();
        var current = new List<PdfTextSpan>();
        double currentY = ordered[0].Y;
        double currentFont = ordered[0].FontSize;
        foreach (var s in ordered) {
            if (current.Count == 0) { current.Add(s); currentY = s.Y; continue; }
            double tolAbs = Math.Min(options.LineMergeMaxPoints, Math.Min(currentFont, s.FontSize) * options.LineMergeToleranceEm);
            if (tolAbs < 0.5) tolAbs = 0.5;
            if (Math.Abs(s.Y - currentY) <= tolAbs) {
                current.Add(s);
                currentFont = (currentFont * (current.Count - 1) + s.FontSize) / current.Count;
            } else {
                lines.Add(BuildLine(current, options));
                current.Clear();
                current.Add(s);
                currentY = s.Y; currentFont = s.FontSize;
            }
        }
        if (current.Count > 0) lines.Add(BuildLine(current, options));
        // Drop obvious duplicate lines drawn twice at the same Y (e.g., shadow/overprint)
        lines = DeduplicateLines(lines);
        return lines;
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
        if (columns.IsTwoColumns) {
            var left = lines.Where(l => l.XStart >= columns.Left.From && l.XStart <= columns.Left.To).OrderByDescending(l => l.Y);
            var right = lines.Where(l => l.XStart >= columns.Right.From && l.XStart <= columns.Right.To).OrderByDescending(l => l.Y);
            bool first = true;
            foreach (var ln in left) { if (!first) sb.Append('\n'); sb.Append(ln.Text); first = false; }
            foreach (var ln in right) { if (!first) sb.Append('\n'); sb.Append(ln.Text); first = false; }
        } else {
            foreach (var ln in lines.OrderByDescending(l => l.Y)) {
                if (sb.Length > 0) sb.Append('\n');
                sb.Append(ln.Text);
            }
        }
        string text = sb.ToString();
        if (options?.JoinHyphenationAcrossLines == true) {
            text = JoinHyphenation(text);
        }
        return text;
    }

    private static string JoinHyphenation(string text) {
        // Collapse word-hyphen-newline-lowercase into wordlowercase
        // Also handle soft hyphen (U+00AD) cases just in case
        return System.Text.RegularExpressions.Regex.Replace(text, "(?<=[A-Za-z])(?:-|\u00AD)\n(?=[a-z])", "");
    }

    private static TextLine BuildLine(List<PdfTextSpan> spans, Options? options) {
        // X sort within the line
        spans.Sort((a, b) => a.X.CompareTo(b.X));
        double xs = spans[0].X;
        var last = spans[spans.Count - 1];
        double xe = last.X + Math.Max(0, last.Advance);
        var text = new StringBuilder();
        for (int i = 0; i < spans.Count; i++) {
            var s = spans[i];
            if (i > 0) {
                // Add a space heuristically if large X gap between spans
                var prev = spans[i - 1];
                double prevEnd = prev.X + Math.Max(0, prev.Advance);
                double gap = s.X - prevEnd;
                // dynamic threshold based on previous span's average glyph advance
                double prevAvg = SafeAvgAdvance(prev);
                double glyphFactor = options?.GapGlyphFactor ?? 0.6;
                double glyphThreshold = prevAvg * glyphFactor;
                // fallback to em threshold when prevAvg unavailable
                double emThreshold = (options?.GapSpaceThresholdEm ?? 0.25) * s.FontSize;
                double threshold = Math.Max(emThreshold, glyphThreshold);
                bool isLeader = IsDotLeader(s.Text);
                bool prevLeader = IsDotLeader(prev.Text);
                // Tight word-join rule: letters adjacent use stricter threshold (slightly more permissive)
                if (IsWordJoin(prev.Text, s.Text)) {
                    // be less aggressive: add space whenever gap exceeds ~0.65x glyph-advance or 0.30em
                    double tight = System.Math.Max(1.0, System.Math.Min(3.0, System.Math.Min(prevAvg * 0.65, s.FontSize * 0.30)));
                    if (gap > tight) text.Append(' ');
                    else {
                        // Fallback: if both look like full words and there is a visible gap, insert a space
                        bool bothAlphaLong = AllWordish(prev.Text) && AllWordish(s.Text) && prev.Text.Length >= 2 && s.Text.Length >= 2;
                        if (bothAlphaLong && gap > 0.8 && (text.Length > 0 && text[text.Length - 1] != ' ')) text.Append(' ');
                    }
                } else if (!isLeader) {
                    // Guard: if both chunks look like full words (>=2 letters) and there is any visible gap, emit a space
                    bool bothAlphaLong = AllWordish(prev.Text) && AllWordish(s.Text) && prev.Text.Length >= 2 && s.Text.Length >= 2;
                    if (bothAlphaLong && gap > 0.8 && (text.Length > 0 && text[text.Length - 1] != ' ')) {
                        text.Append(' ');
                    }
                    if (gap > threshold) text.Append(' ');
                } else {
                    if (gap > 0 && text.Length > 0 && text[text.Length - 1] != ' ') text.Append(' '); // one space before leader
                }
            }
            // drop duplicate shadows: if same text repeats with almost no gap
            if (text.Length > 0 && IsSameAsTail(text, s.Text) && i > 0) {
                var prev = spans[i - 1];
                double prevEnd = prev.X + Math.Max(0, prev.Advance);
                if (s.X - prevEnd < 0.8) {
                    continue;
                }
            }
            text.Append(s.Text);
            // if leader followed by number, ensure a single space
            if (IsDotLeader(s.Text) && i + 1 < spans.Count && IsAllDigits(spans[i + 1].Text)) {
                if (text.Length > 0 && text[text.Length - 1] != ' ') text.Append(' ');
            }
        }
        string outText = text.ToString();
        if (!IsDotLeader(outText)) outText = NormalizeLineText(outText);
        return new TextLine(spans[0].Y, xs, xe, outText, new List<PdfTextSpan>(spans));
    }

    private static double Median(IEnumerable<double> seq) {
        var list = seq.Where(v => v > 0).OrderBy(v => v).ToList();
        if (list.Count == 0) return 12;
        int mid = list.Count / 2;
        if (list.Count % 2 == 1) return list[mid];
        return (list[mid - 1] + list[mid]) / 2.0;
    }

    private static int Clamp(int v, int min, int max) => v < min ? min : (v > max ? max : v);

    private static List<TextLine> DeduplicateLines(List<TextLine> lines) {
        if (lines.Count <= 1) return lines;
        var result = new List<TextLine>(lines.Count);
        var used = new bool[lines.Count];
        for (int i = 0; i < lines.Count; i++) {
            if (used[i]) continue;
            var a = lines[i];
            result.Add(a);
            for (int j = i + 1; j < lines.Count; j++) {
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

    private static bool IsDotLeader(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        for (int i = 0; i < s.Length; i++) if (s[i] != '.') return false; return true;
    }
    private static bool IsWordJoin(string left, string right) {
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)) return false;
        char a = left[left.Length - 1]; char b = right[0];
        bool aWord = char.IsLetterOrDigit(a) || a == ')' || a == '"' || a == '\'';
        bool bWord = char.IsLetterOrDigit(b) || b == '(' || b == '"' || b == '\'';
        return aWord && bWord;
    }
    private static bool IsSameAsTail(StringBuilder sb, string s) {
        if (string.IsNullOrEmpty(s)) return false; int len = s.Length; if (sb.Length < len) return false;
        for (int i = 0; i < len; i++) if (sb[sb.Length - len + i] != s[i]) return false; return true;
    }
    private static bool IsAllDigits(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        for (int i = 0; i < s.Length; i++) if (!char.IsDigit(s[i])) return false; return true;
    }
    private static bool IsWordish(char c) => char.IsLetter(c) || c == '\'' || c == '-' || c == '/';
    private static bool AllWordish(string s) { if (string.IsNullOrEmpty(s)) return false; for (int i = 0; i < s.Length; i++) if (!IsWordish(s[i])) return false; return true; }
    private static bool IsShortAbbrev(string s) { if (string.IsNullOrEmpty(s) || s.Length > 3) return false; for (int i = 0; i < s.Length; i++) if (!char.IsUpper(s[i])) return false; return true; }
    private static string NormalizeLineText(string s) {
        if (string.IsNullOrEmpty(s)) return s;
        s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
        var parts = s.Split(' ');
        if (parts.Length <= 2) {
            if (parts.Length == 2 && AllWordish(parts[0]) && AllWordish(parts[1])) {
                // join two fragments when they are clearly word parts
                if (parts[0].Length == 1 && parts[1].Length >= 3) return parts[0] + parts[1];
                if (parts[1].Length <= 2 || parts[0].Length <= 2) return parts[0] + parts[1];
            }
            return s;
        }
        int shortCount = parts.Count(p => p.Length <= 2 && AllWordish(p));
        if (shortCount < 2) return s;
        var sb = new System.Text.StringBuilder(s.Length);
        sb.Append(parts[0]);
        for (int i = 1; i < parts.Length; i++) {
            string prev = parts[i - 1]; string cur = parts[i];
            bool upperSinglesJoin = prev.Length == 1 && cur.Length == 1 && char.IsUpper(prev[0]) && char.IsUpper(cur[0])
                                    && (i + 1 < parts.Length && parts[i + 1].Length == 1 && char.IsUpper(parts[i + 1][0]));
            bool leadingLetterJoin = AllWordish(prev) && AllWordish(cur) && prev.Length == 1 && cur.Length >= 3;
            bool joinSmall = AllWordish(prev) && AllWordish(cur) && ((prev.Length <= 2 || cur.Length <= 2) || leadingLetterJoin || upperSinglesJoin) && !(IsShortAbbrev(prev) && IsShortAbbrev(cur) && !upperSinglesJoin);
            bool nextShort = (i + 1 < parts.Length) && parts[i + 1].Length <= 2 && AllWordish(parts[i + 1]) && !IsShortAbbrev(parts[i + 1]);
            if (joinSmall || (AllWordish(cur) && cur.Length <= 2 && nextShort)) sb.Append(cur);
            else sb.Append(' ').Append(cur);
        }
        string joined = sb.ToString().Replace("  ", " ");
        // Secondary pass: join common suffix fragments
        var toks = joined.Split(' ');
        if (toks.Length > 1) {
            var sb2 = new System.Text.StringBuilder(joined.Length);
            sb2.Append(toks[0]);
            for (int i = 1; i < toks.Length; i++) {
                string prev = toks[i - 1]; string cur = toks[i];
                if (AllWordish(prev) && AllWordish(cur) && CommonSuffixes.Contains(cur)) sb2.Append(cur);
                else sb2.Append(' ').Append(cur);
            }
            joined = sb2.ToString();
        }
        return joined;
    }
    private static double SafeAvgAdvance(PdfTextSpan span) {
        if (span.Advance <= 0) return span.FontSize * 0.5;
        int len = span.Text?.Length ?? 0; if (len <= 0) return span.FontSize * 0.5;
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
        var engineOpts = options?.ToEngineOptions();
        var lines = TextLayoutEngine.BuildLines(spans, engineOpts);
        var (w, _) = page.GetPageSize();
        // Optional header/footer filtering
        if (options is not null && (options.IgnoreHeaderHeight > 0 || options.IgnoreFooterHeight > 0)) {
            var (_, h) = page.GetPageSize();
            double topCut = h - options.IgnoreHeaderHeight;
            double bottomCut = options.IgnoreFooterHeight;
            lines = lines.Where(l => (options.IgnoreHeaderHeight <= 0 || l.Y < topCut)
                                  && (options.IgnoreFooterHeight <= 0 || l.Y > bottomCut)).ToList();
        }
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
        var engineOpts = options?.ToEngineOptions();
        return ContentStructureExtractor.Extract(spans, engineOpts ?? new TextLayoutEngine.Options());
    }
}
