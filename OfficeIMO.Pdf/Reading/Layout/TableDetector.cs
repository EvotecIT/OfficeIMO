using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Pdf;

/// <summary>
/// Very simple table detector that splits a line into cells when there are large X gaps
/// between adjacent spans. Intended as a first cut for diagnostics and quick CSV-like rows.
/// Heuristics:
/// - Uses per-line span X coordinates and advances to compute inter-span gaps
/// - Splits when gap exceeds max(2*em, 18pt)
/// - Emits a row when at least two cells are produced and one cell is numeric-ish
/// </summary>
internal static class TableDetector {
    public static List<string[]> Detect(List<TextLayoutEngine.TextLine> lines) {
        var rows = new List<string[]>();
        foreach (var ln in lines) {
            if (ln.Spans.Count < 2) continue;
            var cells = SplitByGaps(ln);
            if (cells.Length >= 2 && LooksTabular(cells)) rows.Add(cells);
        }
        return rows;
    }

    public static List<StructuredTable> DetectTablesFromBands(List<List<TextLayoutEngine.TextLine>> bands) {
        var tables = new List<StructuredTable>();
        // Leader-dominated bands should become leader tables, not generic band tables
        foreach (var band in bands) {
            if (band.Count == 0) continue;
            if (IsLeaderBand(band)) {
                var leader = BuildLeaderTableForBand(band);
                if (leader != null && leader.Rows.Count > 0) tables.Add(leader);
            }
        }
        // Then, attempt to form multi-band table groups with similar split positions (non-leader bands only)
        var nonLeaderBands = bands.Where(b => b.Count > 0 && !IsLeaderBand(b)).ToList();
        var grouped = DetectTablesAcrossBandGroups(nonLeaderBands);
        tables.AddRange(grouped);

        // Fallback per-band splits for remaining non-leader bands
        if (tables.Count == 0) {
            foreach (var band in nonLeaderBands) {
                var splits = InferSplits(band);
                if (splits.Count == 0) continue;
                var table = BuildTableFromLinesAndSplits(band, splits, "band-splits");
                if (table != null && table.Rows.Count >= 2) tables.Add(table);
            }
        }
        return tables;
    }

    private static bool IsLeaderBand(List<TextLayoutEngine.TextLine> band) {
        if (band.Count == 0) return false;
        int leaderLines = 0; int nonEmpty = 0;
        foreach (var ln in band) {
            if (string.IsNullOrWhiteSpace(ln.Text)) continue; nonEmpty++;
            if (TryLeaderRowFromLine(ln, out _, out _, out _)) { leaderLines++; continue; }
            bool hasDotSpan = ln.Spans.Any(s => IsDotLeader(s.Text) && s.Text.Length >= 3);
            bool looksLeader = LooksLeaderText(ln.Text);
            if (hasDotSpan || looksLeader) leaderLines++;
        }
        if (nonEmpty == 0) return false;
        // Consider leader band if we have at least 3 leader-like rows, or >=30% of lines
        return leaderLines >= 3 || (leaderLines * 10 >= nonEmpty * 3);
    }

    private static StructuredTable? BuildLeaderTableForBand(List<TextLayoutEngine.TextLine> band) {
        var rows = new List<string[]>();
        double leftMin = double.MaxValue, leftMax = double.MinValue;
        double rightMin = double.MaxValue, rightMax = double.MinValue;
        foreach (var ln in band) {
            if (TryLeaderRowFromLine(ln, out var row, out var left, out var right)) {
                rows.Add(row);
                leftMin = Math.Min(leftMin, left.From);
                leftMax = Math.Max(leftMax, left.To);
                rightMin = Math.Min(rightMin, right.From);
                rightMax = Math.Max(rightMax, right.To);
            }
        }
        if (rows.Count == 0) return null;
        var t = new StructuredTable { YTop = band[0].Y, YBottom = band[band.Count - 1].Y, Kind = "leaders" };
        t.Columns.Add(new StructuredTableColumn { From = leftMin, To = leftMax });
        t.Columns.Add(new StructuredTableColumn { From = rightMin, To = rightMax });
        t.Rows.AddRange(rows);
        return t;
    }

    private static List<StructuredTable> DetectTablesAcrossBandGroups(List<List<TextLayoutEngine.TextLine>> bands) {
        var result = new List<StructuredTable>();
        // Pre-compute splits per band
        var bandSplits = new List<(int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits)>();
        for (int i = 0; i < bands.Count; i++) {
            var b = bands[i]; if (b.Count == 0) continue;
            var sp = InferSplits(b);
            if (sp.Count == 0) continue;
            bandSplits.Add((i, b, sp));
        }
        int k = 0;
        while (k < bandSplits.Count) {
            int start = k;
            var baseSplits = bandSplits[k].splits;
            int end = k;
            // Extend group while splits remain similar and bands are consecutive
            while (end + 1 < bandSplits.Count
                   && bandSplits[end + 1].idx == bandSplits[end].idx + 1
                   && AreSplitsSimilar(baseSplits, bandSplits[end + 1].splits)) {
                end++;
            }
            // Build table for [start..end]
            var groupLines = new List<TextLayoutEngine.TextLine>();
            for (int i = start; i <= end; i++) groupLines.AddRange(bandSplits[i].lines);
            var table = BuildTableFromLinesAndSplits(groupLines, baseSplits, "band-group");
            if (table != null && table.Rows.Count >= 3) result.Add(table);
            k = end + 1;
        }
        return result;
    }

    private static bool AreSplitsSimilar(List<double> a, List<double> b) {
        if (a.Count != b.Count) return false;
        double tol = 16.0; // points
        for (int i = 0; i < a.Count; i++) if (Math.Abs(a[i] - b[i]) > tol) return false;
        return true;
    }

    private static StructuredTable? BuildTableFromLinesAndSplits(List<TextLayoutEngine.TextLine> lines, List<double> splits, string kind) {
        if (splits.Count == 0) return null;
        double minX = double.MaxValue, maxX = double.MinValue;
        foreach (var ln in lines) { minX = Math.Min(minX, ln.XStart); maxX = Math.Max(maxX, ln.XEnd); }
        var table = new StructuredTable {
            YTop = lines[0].Y,
            YBottom = lines[lines.Count - 1].Y,
            Kind = kind
        };
        double prev = minX;
        for (int i = 0; i <= splits.Count; i++) {
            double next = (i < splits.Count) ? splits[i] : maxX;
            table.Columns.Add(new StructuredTableColumn { From = prev, To = next });
            prev = next;
        }
        int cols = table.Columns.Count;
        foreach (var ln in lines) {
            var cells = SplitBySplits(ln, splits);
            if (cells.Length != cols) continue;
            bool anyContent = false; for (int i = 0; i < cells.Length; i++) if (!string.IsNullOrWhiteSpace(cells[i])) { anyContent = true; break; }
            if (!anyContent) continue;
            table.Rows.Add(cells);
        }
        return table.Rows.Count > 0 ? table : null;
    }

    public static StructuredTable? DetectLeaderTable(List<TextLayoutEngine.TextLine> lines) {
        var candidates = lines.Where(l => !string.IsNullOrWhiteSpace(l.Text)).ToList();
        if (candidates.Count == 0) return null;
        var rows = new List<string[]>();
        double leftMin = double.MaxValue, leftMax = double.MinValue;
        double rightMin = double.MaxValue, rightMax = double.MinValue;
        foreach (var ln in candidates) {
            if (TryLeaderRowFromLine(ln, out var row, out var leftBounds, out var rightBounds)) {
                rows.Add(row);
                leftMin = Math.Min(leftMin, leftBounds.From);
                leftMax = Math.Max(leftMax, leftBounds.To);
                rightMin = Math.Min(rightMin, rightBounds.From);
                rightMax = Math.Max(rightMax, rightBounds.To);
            }
        }
        if (rows.Count < 2) return null;
        var table = new StructuredTable {
            YTop = candidates[0].Y,
            YBottom = candidates[candidates.Count - 1].Y,
            Kind = "leaders"
        };
        table.Columns.Add(new StructuredTableColumn { From = leftMin, To = leftMax });
        table.Columns.Add(new StructuredTableColumn { From = rightMin, To = rightMax });
        table.Rows.AddRange(rows);
        return table;
    }

    /// <summary>
    /// Band-aware detection that first infers stable column split positions within each band,
    /// then splits lines consistently into those columns.
    /// </summary>
    public static List<string[]> DetectFromBands(List<List<TextLayoutEngine.TextLine>> bands) {
        var all = new List<string[]>();
        foreach (var band in bands) {
            if (band.Count == 0) continue;
            var splits = InferSplits(band);
            if (splits.Count == 0) {
                // fallback to per-line gap splitting
                foreach (var ln in band) {
                    if (ln.Spans.Count < 2) continue;
                    var cells = SplitByGaps(ln);
                    if (cells.Length >= 2 && LooksTabular(cells)) all.Add(cells);
                }
                continue;
            }
            // Consistent splitting using inferred splits
            int cols = splits.Count + 1;
            foreach (var ln in band) {
                var cells = SplitBySplits(ln, splits);
                if (cells.Length == cols) {
                    bool any = false; for (int i = 0; i < cells.Length; i++) if (!string.IsNullOrWhiteSpace(cells[i])) { any = true; break; }
                    if (any) all.Add(cells);
                }
            }
        }
        return all;
    }

    private static List<double> InferSplits(List<TextLayoutEngine.TextLine> lines) {
        // Collect candidate split X positions as midpoints of large gaps between adjacent spans
        var cands = new List<double>();
        int eligibleLines = 0;
        foreach (var ln in lines) {
            if (ln.Spans.Count < 2) continue;
            eligibleLines++;
            // Dot-leader spans are strong split hints
            for (int k = 0; k < ln.Spans.Count; k++) {
                var s = ln.Spans[k];
                if (IsDotLeader(s.Text)) {
                    double mid = s.X + Math.Max(0, s.Advance) / 2.0;
                    cands.Add(mid);
                }
            }
            for (int i = 1; i < ln.Spans.Count; i++) {
                var prev = ln.Spans[i - 1]; var curSpan = ln.Spans[i];
                double prevEnd = prev.X + Math.Max(0, prev.Advance);
                double gap = curSpan.X - prevEnd;
                double em = Math.Max(prev.FontSize, curSpan.FontSize);
                double threshold = Math.Max(18.0, em * 2.0);
                if (gap >= threshold) {
                    double mid = prevEnd + (gap / 2.0);
                    cands.Add(mid);
                }
            }
        }
        if (eligibleLines == 0 || cands.Count == 0) return new List<double>();
        // Histogram candidates into 4pt bins and select peaks with sufficient votes
        double binW = 4.0;
        double minX = cands.Min(); double maxX = cands.Max();
        int bins = Math.Max(1, (int)Math.Ceiling((maxX - minX) / binW));
        var hist = new int[bins];
        foreach (var x in cands) {
            int b = (int)Math.Floor((x - minX) / binW);
            if (b < 0) b = 0; if (b >= bins) b = bins - 1; hist[b]++;
        }
        int voteCut = Math.Max(2, (int)Math.Ceiling(eligibleLines * 0.35));
        var peaks = new List<double>();
        for (int b = 0; b < bins; b++) if (hist[b] >= voteCut) peaks.Add(minX + b * binW + binW / 2.0);
        if (peaks.Count == 0) {
            // Fallback for narrow bands: pick the strongest bin if any votes exist
            int maxVotes = 0; int maxBin = -1;
            for (int b = 0; b < bins; b++) if (hist[b] > maxVotes) { maxVotes = hist[b]; maxBin = b; }
            if (maxVotes > 0 && maxBin >= 0) peaks.Add(minX + maxBin * binW + binW / 2.0);
            else return new List<double>();
        }
        // Merge nearby peaks (< 16pt apart)
        peaks.Sort();
        var merged = new List<double>();
        double acc = peaks[0]; int count = 1;
        for (int i = 1; i < peaks.Count; i++) {
            if (Math.Abs(peaks[i] - acc) < 16.0) { acc = (acc * count + peaks[i]) / (count + 1); count++; }
            else { merged.Add(acc); acc = peaks[i]; count = 1; }
        }
        merged.Add(acc);
        // Limit to a reasonable number of splits to avoid over-fragmentation
        if (merged.Count > 6) merged = merged.Take(6).ToList();
        return merged;
    }

    private static string[] SplitBySplits(TextLayoutEngine.TextLine ln, List<double> splits) {
        int cols = splits.Count + 1;
        var cellBuilders = new System.Text.StringBuilder[cols];
        for (int i = 0; i < cols; i++) cellBuilders[i] = new System.Text.StringBuilder();
        int ColIndex(double x) { int idx = 0; while (idx < splits.Count && x >= splits[idx]) idx++; return idx; }
        for (int i = 0; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            int cidx = ColIndex(s.X);
            var sb = cellBuilders[cidx];
            if (sb.Length > 0 && !char.IsWhiteSpace(sb[sb.Length - 1])) sb.Append(' ');
            sb.Append(s.Text);
        }
        var cells = new string[cols];
        for (int i = 0; i < cols; i++) cells[i] = cellBuilders[i].ToString().Trim();
        return cells;
    }

    private static string[] SplitByGaps(TextLayoutEngine.TextLine ln) {
        // Determine gaps between spans using XEnd(prev) -> XStart(next)
        double ThresholdFor(PdfTextSpan prev, PdfTextSpan next) {
            double em = Math.Max(prev.FontSize, next.FontSize);
            return Math.Max(18.0, em * 2.0); // 18pt or 2em
        }
        var cells = new List<string>();
        var current = new System.Text.StringBuilder();
        for (int i = 0; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            if (i > 0) {
                var p = ln.Spans[i - 1];
                double prevEnd = p.X + Math.Max(0, p.Advance);
                double gap = s.X - prevEnd;
                if (gap > ThresholdFor(p, s)) {
                    // split to a new cell
                    cells.Add(current.ToString().Trim());
                    current.Clear();
                } else if (gap > 1.0 && (current.Length > 0 && current[current.Length - 1] != ' ')) {
                    // small gap -> ensure single space
                    current.Append(' ');
                }
            }
            current.Append(s.Text);
        }
        if (current.Length > 0) cells.Add(current.ToString().Trim());
        return cells.ToArray();
    }

    private static bool LooksTabular(string[] cells) {
        // Require at least one numeric-ish cell and avoid one-word rows
        bool anyNumeric = cells.Any(c => HasManyDigits(c));
        bool hasContent = cells.Any(c => c.Length >= 2);
        return anyNumeric && hasContent;
    }

    private static bool HasManyDigits(string s) {
        int digits = 0; for (int i = 0; i < s.Length; i++) if (char.IsDigit(s[i])) digits++;
        return digits >= Math.Max(2, s.Length / 4);
    }

    private static bool IsDotLeader(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        for (int i = 0; i < s.Length; i++) if (s[i] != '.') return false; return true;
    }

    private static bool LooksLeaderText(string s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        int dots = 0; for (int i = 0; i < s.Length; i++) if (s[i] == '.') dots++;
        return dots >= 3;
    }

    private static bool TryLeaderRowFromLine(TextLayoutEngine.TextLine ln, out string[] row, out (double From,double To) left, out (double From,double To) right) {
        row = Array.Empty<string>(); left = (0,0); right=(0,0);
        // Find a dotted leader span in this line
        int leaderIdx = -1;
        for (int i = 0; i < ln.Spans.Count; i++) if (IsDotLeader(ln.Spans[i].Text) && ln.Spans[i].Text.Length >= 3) { leaderIdx = i; break; }
        if (leaderIdx < 0) return false;
        // Left label: join spans before leader (preserve minimal spaces)
        var sbLeft = new System.Text.StringBuilder();
        double leftFrom = double.MaxValue, leftTo = double.MinValue;
        for (int i = 0; i < leaderIdx; i++) {
            var s = ln.Spans[i];
            if (sbLeft.Length > 0) sbLeft.Append(' ');
            sbLeft.Append(s.Text);
            leftFrom = Math.Min(leftFrom, s.X);
            leftTo = Math.Max(leftTo, s.X + Math.Max(0, s.Advance));
        }
        string leftText = CleanLeftLabel(sbLeft.ToString());
        // Right number: consume digits after leader
        var sbRight = new System.Text.StringBuilder();
        double rightFrom = double.MaxValue, rightTo = double.MinValue;
        for (int i = leaderIdx + 1; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            if (ContainsDigit(s.Text)) {
                if (sbRight.Length > 0 && sbRight[sbRight.Length - 1] != ' ') sbRight.Append(' ');
                sbRight.Append(s.Text);
                rightFrom = Math.Min(rightFrom, s.X);
                rightTo = Math.Max(rightTo, s.X + Math.Max(0, s.Advance));
            }
        }
        string rightText = sbRight.ToString().Trim();
        // Sanity checks
        if (leftText.Length == 0 || rightText.Length == 0) return false;
        // Strip non-digits and spaces from right
        rightText = new string(rightText.Where(char.IsDigit).ToArray());
        if (rightText.Length == 0) return false;
        row = new [] { leftText, rightText };
        left = (leftFrom, leftTo);
        right = (rightFrom, rightTo);
        return true;
    }

    private static bool ContainsDigit(string s) { for (int i = 0; i < s.Length; i++) if (char.IsDigit(s[i])) return true; return false; }

        private static string CleanLeftLabel(string s) {
            if (string.IsNullOrEmpty(s)) return s;
            // Normalize spaces
            s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
            // Remove trailing dots on label (leaders)
            s = s.Trim('.');
            // Remove repeated dot groups inside label
            s = System.Text.RegularExpressions.Regex.Replace(s, "[.]{2,}", ".");
            // Tidy quotes and parentheses spacing
            s = s.Replace(" ' ", " '").Replace("( ", "(").Replace(" )", ")");
            // Re-insert spaces around common glued prepositions if camel-cased inside
            s = System.Text.RegularExpressions.Regex.Replace(s, "([A-Za-z])of([A-Z])", "$1 of $2");
            s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z]{2,})of([A-Z])", "$1 of $2");
            s = System.Text.RegularExpressions.Regex.Replace(s, "([A-Za-z])in([A-Z])", "$1 in $2");
            s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z]{2,})in([A-Z])", "$1 in $2");
            s = System.Text.RegularExpressions.Regex.Replace(s, "([A-Za-z])and([A-Z])", "$1 and $2");
            s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z]{2,})and([A-Z])", "$1 and $2");
            // generic lower→Upper split (camel-case → spaced)
            s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z])([A-Z])", "$1 $2");
            // Collapse micro-token shattering (aggressive but safe-ish for leaders)
            var parts = s.Split(' ');
        if (parts.Length <= 2) return s;
        bool Wordish(string t) { for (int i = 0; i < t.Length; i++) { char c = t[i]; if (!(char.IsLetterOrDigit(c) || c=='\''||c=='-'||c=='/')) return false; } return t.Length>0; }
        bool ShortAbbrev(string t) { if (t.Length==0 || t.Length>3) return false; for (int i=0;i<t.Length;i++) if(!char.IsUpper(t[i])) return false; return true; }
        int shortCount = parts.Count(p => p.Length <= 2 && Wordish(p));
        if (!(shortCount >= 2 || shortCount * 4 >= parts.Length)) return s;
        var sb = new System.Text.StringBuilder(s.Length);
        sb.Append(parts[0]);
        for (int i = 1; i < parts.Length; i++) {
            string prev = parts[i-1]; string cur = parts[i];
            bool joinSmall = Wordish(prev) && Wordish(cur) && !ShortAbbrev(prev) && !ShortAbbrev(cur) && (prev.Length<=2 || cur.Length<=2);
            bool nextShort = (i+1<parts.Length) && parts[i+1].Length<=2 && Wordish(parts[i+1]) && !ShortAbbrev(parts[i+1]);
            if (joinSmall || (Wordish(cur)&&cur.Length<=2 && nextShort)) sb.Append(cur);
            else sb.Append(' ').Append(cur);
        }
        return sb.ToString().Replace("  ", " ");
    }
}
