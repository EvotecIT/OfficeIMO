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
    public sealed class Options {
        /// <summary>Assume page margins (points) when inferring columns. Default: 36 pt (0.5").</summary>
        public double MarginLeft { get; set; } = 36;
        public double MarginRight { get; set; } = 36;
        /// <summary>Histogram bin width for gutter detection. Default: 5 pt.</summary>
        public double BinWidth { get; set; } = 5;
        /// <summary>Minimum gutter width to consider split into two columns. Default: 24 pt.</summary>
        public double MinGutterWidth { get; set; } = 24;
        /// <summary>Maximum Y delta (as fraction of font size) to group spans into the same line. Default: 1.8.</summary>
        public double LineMergeToleranceEm { get; set; } = 1.8;
        /// <summary>Force single column when true; skip gutter detection.</summary>
        public bool ForceSingleColumn { get; set; } = false;
        /// <summary>Threshold in em units to insert a space between adjacent spans on the same line. Default: 0.3.</summary>
        public double GapSpaceThresholdEm { get; set; } = 0.25;
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
        foreach (var s in ordered) {
            if (current.Count == 0) { current.Add(s); currentY = s.Y; continue; }
            double tolAbs = Math.Max(1.0, Math.Max(medianSize, s.FontSize) * options.LineMergeToleranceEm);
            if (Math.Abs(s.Y - currentY) <= tolAbs) {
                current.Add(s);
                // adapt running baseline to mitigate tiny per-glyph Y drift
                currentY = (currentY * (current.Count - 1) + s.Y) / current.Count;
            } else {
                lines.Add(BuildLine(current, options));
                current.Clear();
                current.Add(s);
                currentY = s.Y;
            }
        }
        if (current.Count > 0) lines.Add(BuildLine(current, options));
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
        double xe = last.X + ApproxWidth(last);
        var text = new StringBuilder();
        for (int i = 0; i < spans.Count; i++) {
            var s = spans[i];
            if (i > 0) {
                // Add a space heuristically if large X gap between spans
                double prevEnd = spans[i - 1].X + ApproxWidth(spans[i - 1]);
                double gapEm = options?.GapSpaceThresholdEm ?? 0.3;
                if (s.X - prevEnd > s.FontSize * gapEm) text.Append(' ');
            }
            text.Append(s.Text);
        }
        return new TextLine(spans[0].Y, xs, xe, text.ToString(), new List<PdfTextSpan>(spans));
    }

    private static double ApproxWidth(PdfTextSpan s) {
        // A crude approximation: average width ~ 0.5–0.6 em per glyph
        double em = 0.55;
        return Math.Max(0, s.Text?.Length ?? 0) * s.FontSize * em;
    }

    private static double Median(IEnumerable<double> seq) {
        var list = seq.Where(v => v > 0).OrderBy(v => v).ToList();
        if (list.Count == 0) return 12;
        int mid = list.Count / 2;
        if (list.Count % 2 == 1) return list[mid];
        return (list[mid - 1] + list[mid]) / 2.0;
    }

    private static int Clamp(int v, int min, int max) => v < min ? min : (v > max ? max : v);
}

/// <summary>
/// Convenience helpers for callers to get column-aware text from a PdfReadPage.
/// </summary>
public static class PdfReadPageExtensions {
    /// <summary>
    /// Extracts text from a page with simple two-column detection when present.
    /// </summary>
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
}
