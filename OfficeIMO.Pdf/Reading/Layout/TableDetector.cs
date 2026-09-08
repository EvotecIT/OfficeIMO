using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Pdf;

/// <summary>
/// Recovers tables from ruled regions, tagged structure, leaders, and stable text geometry.
/// Text-only evidence is based on alignment, spacing, font-relative compactness, and Unicode
/// character properties; it does not depend on language-specific vocabulary or word boundaries.
/// </summary>
internal static partial class TableDetector {
    private const int MaximumPositionedRecoveryLines = 4096;
    private const int MaximumPositionedRecoveryColumns = 64;
    private const int MaximumPositionedRecoveryCells = 65536;
    private const double MaximumCompactCellWidthInFontSizes = 24D;
    private const double MaximumAverageCompactCellWidthInFontSizes = 12D;
    public static List<string[]> Detect(List<TextLayoutEngine.TextLine> lines, double? pageHeight = null) {
        var rows = new List<string[]>();
        foreach (var match in DetectLineRows(lines, pageHeight)) {
            rows.Add(match.Cells);
        }
        return rows;
    }

    public static List<(TextLayoutEngine.TextLine Line, string[] Cells)> DetectLineRows(
        List<TextLayoutEngine.TextLine> lines,
        double? pageHeight = null,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        var rows = new List<(TextLayoutEngine.TextLine Line, string[] Cells)>();
        foreach (var ln in lines) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (!CanRecoverTableLine(ln, pageHeight) || ln.Spans.Count < 2) continue;
            var cells = SplitByGaps(ln);
            if (cells.Length >= 2 && LooksTabular(cells)) rows.Add((ln, cells));
        }
        return rows;
    }

    public static List<StructuredTable> DetectTablesFromBands(
        List<List<TextLayoutEngine.TextLine>> bands,
        double? pageHeight = null,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        var recoverableBands = new List<List<TextLayoutEngine.TextLine>>(bands.Count);
        for (int bandIndex = 0; bandIndex < bands.Count; bandIndex++) {
            cancellationCheck?.Invoke();
            List<TextLayoutEngine.TextLine> sourceBand = bands[bandIndex];
            if (sourceBand.Count > 0) consumeWork?.Invoke(sourceBand.Count);
            List<TextLayoutEngine.TextLine> recoverable = sourceBand
                .Where(line => CanRecoverTableLine(line, pageHeight))
                .ToList();
            if (recoverable.Count > 0) recoverableBands.Add(recoverable);
        }
        bands = recoverableBands;
        var tables = new List<StructuredTable>();
        // Leader-dominated bands should become leader tables, not generic band tables
        foreach (var band in bands) {
            cancellationCheck?.Invoke();
            if (band.Count == 0) continue;
            consumeWork?.Invoke(band.Count);
            if (IsLeaderBand(band)) {
                var leader = BuildLeaderTableForBand(band);
                if (leader != null && leader.Rows.Count > 0) tables.Add(leader);
            }
        }
        // Then, attempt to form multi-band table groups with similar split positions (non-leader bands only)
        var nonLeaderBands = bands.Where(b => b.Count > 0 && !IsLeaderBand(b)).ToList();
        var grouped = DetectTablesAcrossBandGroups(nonLeaderBands, consumeWork, cancellationCheck);
        tables.AddRange(grouped);

        // Fallback per-band splits for remaining non-leader bands
        if (tables.Count == 0) {
            foreach (var band in nonLeaderBands) {
                cancellationCheck?.Invoke();
                if (band.Count > 0) consumeWork?.Invoke(band.Count);
                var splits = InferSplits(band);
                if (splits.Count == 0) continue;
                var table = BuildTableFromLinesAndSplits(band, splits, "band-splits");
                if (table != null && table.Rows.Count >= 2 && HasValidatedRows(table, band)) tables.Add(table);
            }
        }
        List<TextLayoutEngine.TextLine> unmatchedLines = nonLeaderBands
            .SelectMany(static band => band)
            .Where(line => !IsCoveredByDetectedTable(line, tables))
            .Take(MaximumPositionedRecoveryLines)
            .ToList();
        List<StructuredTable> positionedTables = DetectPositionedCellTables(
            unmatchedLines,
            pageHeight,
            consumeWork,
            cancellationCheck);
        if (positionedTables.Count > 0) {
            tables.RemoveAll(table =>
                table.Rows.Count < 3 &&
                !string.Equals(table.Kind, "leaders", StringComparison.Ordinal) &&
                positionedTables.Any(positioned => IsSubsumedByPositionedTable(table, positioned)));
            tables.AddRange(positionedTables);
        }
        return tables;
    }

    private static bool IsSubsumedByPositionedTable(StructuredTable candidate, StructuredTable positioned) {
        if (candidate.Columns.Count == 0 || positioned.Columns.Count == 0 || positioned.Rows.Count <= candidate.Rows.Count) {
            return false;
        }

        double candidateLeft = candidate.Columns.Min(static column => Math.Min(column.From, column.To));
        double candidateRight = candidate.Columns.Max(static column => Math.Max(column.From, column.To));
        double positionedLeft = positioned.Columns.Min(static column => Math.Min(column.From, column.To));
        double positionedRight = positioned.Columns.Max(static column => Math.Max(column.From, column.To));
        double horizontalOverlap = Math.Max(0D, Math.Min(candidateRight, positionedRight) - Math.Max(candidateLeft, positionedLeft));
        double candidateWidth = candidateRight - candidateLeft;
        if (candidateWidth <= 0.001D || horizontalOverlap + 0.001D < candidateWidth * 0.5D) return false;

        double candidateTop = Math.Max(candidate.YTop, candidate.YBottom);
        double candidateBottom = Math.Min(candidate.YTop, candidate.YBottom);
        double positionedTop = Math.Max(positioned.YTop, positioned.YBottom);
        double positionedBottom = Math.Min(positioned.YTop, positioned.YBottom);
        double verticalOverlap = Math.Max(0D, Math.Min(candidateTop, positionedTop) - Math.Max(candidateBottom, positionedBottom));
        double candidateHeight = candidateTop - candidateBottom;
        if (candidateHeight > 0.001D && verticalOverlap + 0.001D < candidateHeight * 0.5D) return false;

        var positionedCells = new HashSet<string>(
            positioned.Rows.SelectMany(static row => row).Where(static cell => !string.IsNullOrWhiteSpace(cell)),
            StringComparer.Ordinal);
        return candidate.Rows
            .SelectMany(static row => row)
            .Where(static cell => !string.IsNullOrWhiteSpace(cell))
            .All(positionedCells.Contains);
    }

    private static bool IsCoveredByDetectedTable(
        TextLayoutEngine.TextLine line,
        List<StructuredTable> tables) {
        for (int index = 0; index < tables.Count; index++) {
            // Two-row band candidates are deliberately admitted only with strong
            // evidence, but they are still too weak to own the source geometry.
            // Let the independent positioned-cell pass inspect those lines so it
            // can recover a complete header/body region or a side-by-side table.
            if (tables[index].Rows.Count < 3 &&
                !string.Equals(tables[index].Kind, "leaders", StringComparison.Ordinal)) {
                continue;
            }
            double top = Math.Max(tables[index].YTop, tables[index].YBottom);
            double bottom = Math.Min(tables[index].YTop, tables[index].YBottom);
            if (line.Y > top + 0.001D || line.Y < bottom - 0.001D || tables[index].Columns.Count == 0) {
                continue;
            }

            double left = tables[index].Columns.Min(static column => Math.Min(column.From, column.To));
            double right = tables[index].Columns.Max(static column => Math.Max(column.From, column.To));
            double lineLeft = Math.Min(line.XStart, line.XEnd);
            double lineRight = Math.Max(line.XStart, line.XEnd);
            double overlap = Math.Max(0D, Math.Min(lineRight, right) - Math.Max(lineLeft, left));
            double narrowerWidth = Math.Min(lineRight - lineLeft, right - left);
            if (narrowerWidth > 0.001D && overlap + 0.001D >= narrowerWidth * 0.5D) return true;
        }
        return false;
    }

    private static bool IsLeaderBand(List<TextLayoutEngine.TextLine> band) {
        if (band.Count == 0) return false;
        int leaderLines = 0; int nonEmpty = 0;
        foreach (var ln in band) {
            if (string.IsNullOrWhiteSpace(ln.Text)) continue; nonEmpty++;
            if (TryLeaderRowFromLine(ln, out _, out _, out _)) { leaderLines++; continue; }
            bool hasLeaderSpan = ln.Spans.Any(s => IsLeaderSpan(s.Text) && s.Text.Length >= 3);
            bool looksLeader = LooksLeaderText(ln.Text);
            if (hasLeaderSpan || looksLeader) leaderLines++;
        }
        if (nonEmpty == 0) return false;
        // Consider leader band if we have at least 3 leader-like rows, or >=30% of lines
        return leaderLines >= 3 || (leaderLines * 10 >= nonEmpty * 3);
    }

    private static StructuredTable? BuildLeaderTableForBand(List<TextLayoutEngine.TextLine> band) {
        var rows = new List<string[]>();
        var sourceRuns = new List<PdfTextSpan>();
        double leftMin = double.MaxValue, leftMax = double.MinValue;
        double rightMin = double.MaxValue, rightMax = double.MinValue;
        foreach (var ln in band) {
            if (TryLeaderRowFromLine(ln, out var row, out var left, out var right)) {
                rows.Add(row);
                sourceRuns.AddRange(ln.Spans);
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
        t.SourceRuns = sourceRuns.Distinct().ToArray();
        return t;
    }

    private static bool AreSplitsSimilar(List<double> a, List<double> b) {
        if (a.Count != b.Count) return false;
        double tol = 16.0; // points
        for (int i = 0; i < a.Count; i++) if (Math.Abs(a[i] - b[i]) > tol) return false;
        return true;
    }

    private static StructuredTable? BuildTableFromLinesAndSplits(
        List<TextLayoutEngine.TextLine> lines,
        List<double> splits,
        string kind,
        Dictionary<TextLayoutEngine.TextLine, List<double>>? lineSplitOverrides = null) {
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
        var sourceRuns = new List<PdfTextSpan>();
        var sourceLines = new List<TextLayoutEngine.TextLine>(lines.Count);
        foreach (var ln in lines) {
            List<double> lineSplits = lineSplitOverrides is not null &&
                                      lineSplitOverrides.TryGetValue(ln, out List<double>? splitOverride)
                ? splitOverride
                : splits;
            var cells = SplitBySplits(ln, lineSplits);
            if (cells.Length != cols) continue;
            bool anyContent = false; for (int i = 0; i < cells.Length; i++) if (!string.IsNullOrWhiteSpace(cells[i])) { anyContent = true; break; }
            if (!anyContent) continue;
            table.Rows.Add(cells);
            sourceRuns.AddRange(ln.Spans);
            sourceLines.Add(ln);
        }
        table.SourceRuns = sourceRuns.Distinct().ToArray();
        table.SourceLines = sourceLines;
        return table.Rows.Count > 0 ? table : null;
    }

    public static StructuredTable? DetectLeaderTable(
        List<TextLayoutEngine.TextLine> lines,
        double? pageHeight = null,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        cancellationCheck?.Invoke();
        if (lines.Count == 0) return null;
        consumeWork?.Invoke(lines.Count);
        var candidates = lines
            .Where(line => CanRecoverTableLine(line, pageHeight))
            .Where(static line => !string.IsNullOrWhiteSpace(line.Text))
            .ToList();
        if (candidates.Count == 0) return null;
        var rows = new List<string[]>();
        var sourceRuns = new List<PdfTextSpan>();
        double leftMin = double.MaxValue, leftMax = double.MinValue;
        double rightMin = double.MaxValue, rightMax = double.MinValue;
        foreach (var ln in candidates) {
            cancellationCheck?.Invoke();
            if (TryLeaderRowFromLine(ln, out var row, out var leftBounds, out var rightBounds)) {
                rows.Add(row);
                sourceRuns.AddRange(ln.Spans);
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
        table.SourceRuns = sourceRuns.Distinct().ToArray();
        return table;
    }

    /// <summary>
    /// Band-aware detection that first infers stable column split positions within each band,
    /// then splits lines consistently into those columns.
    /// </summary>
    public static List<string[]> DetectFromBands(
        List<List<TextLayoutEngine.TextLine>> bands,
        double? pageHeight = null) {
        var all = new List<string[]>();
        foreach (var band in bands) {
            List<TextLayoutEngine.TextLine> recoverable = band
                .Where(line => CanRecoverTableLine(line, pageHeight))
                .ToList();
            if (recoverable.Count == 0) continue;
            var splits = InferSplits(recoverable);
            if (splits.Count == 0) {
                // fallback to per-line gap splitting
                foreach (var ln in recoverable) {
                    if (ln.Spans.Count < 2) continue;
                    var cells = SplitByGaps(ln);
                    if (cells.Length >= 2 && LooksTabular(cells)) all.Add(cells);
                }
                continue;
            }
            // Consistent splitting using inferred splits
            int cols = splits.Count + 1;
            foreach (var ln in recoverable) {
                var cells = SplitBySplits(ln, splits);
                if (cells.Length == cols) {
                    bool any = false; for (int i = 0; i < cells.Length; i++) if (!string.IsNullOrWhiteSpace(cells[i])) { any = true; break; }
                    if (any) all.Add(cells);
                }
            }
        }
        return all;
    }

    private static bool CanRecoverTableLine(TextLayoutEngine.TextLine line, double? pageHeight) =>
        line.Spans.Count > 0 &&
        line.Spans.All(span => span.CanProjectCompleteText(pageHeight));

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
                // A leader is a repeated punctuation run. A single hyphen, dot, or underscore
                // is ordinary cell content (identifiers, decimals, and names) and cannot be
                // treated as a column boundary.
                if (s.Text.Length >= 3 && IsLeaderSpan(s.Text)) {
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
        int voteCut = eligibleLines == 1 ? 1 : Math.Max(2, (int)Math.Ceiling(eligibleLines * 0.35));
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

    private static bool LooksTabular(string[] cells) {
        // Require at least one numeric-ish cell and avoid one-word rows
        bool anyNumeric = cells.Any(c => HasManyDigits(c));
        bool hasContent = cells.Any(c => c.Length >= 2);
        return anyNumeric && hasContent;
    }

    private static bool HasManyDigits(string s) {
        int digits = PdfUnicodeScalarAnalysis.CountDecimalDigits(s);
        int scalars = PdfUnicodeScalarAnalysis.CountScalars(s);
        return digits >= Math.Max(2, scalars / 4);
    }

    private static bool IsLeaderSpan(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        char c = s[0];
        if (c != '.' && c != '-' && c != '_') return false;
        for (int i = 1; i < s.Length; i++) if (s[i] != c) return false; return true;
    }

    private static bool LooksLeaderText(string s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        char previous = '\0';
        int runLength = 0;
        for (int i = 0; i < s.Length; i++) {
            char current = s[i];
            if (current != '.' && current != '-' && current != '_') {
                previous = '\0';
                runLength = 0;
                continue;
            }

            runLength = current == previous ? runLength + 1 : 1;
            if (runLength >= 3) return true;
            previous = current;
        }
        return false;
    }

    private static bool TryLeaderRowFromLine(TextLayoutEngine.TextLine ln, out string[] row, out (double From,double To) left, out (double From,double To) right) {
        row = Array.Empty<string>(); left = (0,0); right=(0,0);
        // Find a leader span in this line
        int leaderIdx = -1;
        for (int i = 0; i < ln.Spans.Count; i++) if (IsLeaderSpan(ln.Spans[i].Text) && ln.Spans[i].Text.Length >= 3) { leaderIdx = i; break; }
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
        // Right value: consume the value spans after leader, preserving numeric punctuation.
        var sbRight = new System.Text.StringBuilder();
        double rightFrom = double.MaxValue, rightTo = double.MinValue;
        for (int i = leaderIdx + 1; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            if (IsLeaderSpan(s.Text)) {
                continue;
            }

            if (sbRight.Length > 0 && sbRight[sbRight.Length - 1] != ' ') sbRight.Append(' ');
            sbRight.Append(s.Text);
            rightFrom = Math.Min(rightFrom, s.X);
            rightTo = Math.Max(rightTo, s.X + Math.Max(0, s.Advance));
        }
        string rightText = NormalizeLeaderValue(sbRight.ToString());
        // Sanity checks
        if (leftText.Length == 0 || rightText.Length == 0) return false;
        row = new [] { leftText, rightText };
        left = (leftFrom, leftTo);
        right = (rightFrom, rightTo);
        return true;
    }

    private static string CleanLeftLabel(string s) {
        if (string.IsNullOrEmpty(s)) return s;
        return System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
    }

    private static string NormalizeLeaderValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        string normalized = System.Text.RegularExpressions.Regex.Replace(value.Trim(), "\\s+", " ");

        return PdfUnicodeScalarAnalysis.ContainsDecimalDigit(normalized) ? normalized : string.Empty;
    }
}
