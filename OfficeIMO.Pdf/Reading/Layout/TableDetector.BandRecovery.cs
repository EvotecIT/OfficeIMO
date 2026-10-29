namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    private static List<PdfStructuredTable> DetectTablesAcrossBandGroups(
        List<List<TextLayoutEngine.TextLine>> bands,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var result = new List<PdfStructuredTable>();
        // Pre-compute splits per band
        var bandSplits = new List<(int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits)>();
        for (int i = 0; i < bands.Count; i++) {
            cancellationCheck?.Invoke();
            if (bands[i].Count > 0) consumeWork?.Invoke(bands[i].Count);
            var b = bands[i]; if (b.Count == 0) continue;
            var sp = InferSplits(b);
            if (sp.Count == 0) continue;
            bandSplits.Add((i, b, sp));
        }
        int k = 0;
        var claimedBandIndexes = new HashSet<int>();
        while (k < bandSplits.Count) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            int start = k;
            var baseSplits = bandSplits[k].splits;
            int end = k;
            var includedBridgeBandIndexes = new HashSet<int>();
            bool requiresAlignedCellSplits = false;
            var alignedSplitAccumulator = new AlignedSplitAccumulator(
                baseSplits.Count + 1,
                bandSplits[start].lines);
            int precedingBandIndex = bandSplits[start].idx - 1;
            List<TextLayoutEngine.TextLine>? headerLines = claimedBandIndexes.Contains(precedingBandIndex)
                ? null
                : TryGetPrecedingHeaderLines(
                    bands,
                    bandSplits[start].idx,
                    baseSplits);
            // Extend while splits remain similar. An intervening band without
            // detectable splits is either retained as a strongly evidenced
            // spanning row or skipped when it belongs to an adjacent region.
            while (end + 1 < bandSplits.Count) {
                (int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits) current = bandSplits[end];
                (int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits) next = bandSplits[end + 1];
                List<TextLayoutEngine.TextLine>? previous = end > start
                    ? bandSplits[end - 1].lines
                    : headerLines;
                List<TextLayoutEngine.TextLine>? following = end + 2 < bandSplits.Count &&
                                                              bandSplits[end + 2].idx == next.idx + 1
                    ? bandSplits[end + 2].lines
                    : null;
                if ((end > start || headerLines is not null) &&
                    StartsStructurallySeparatedTable(current.lines, next.lines, previous, following)) {
                    break;
                }

                bool hasNonLeftAlignedCells = BandsHaveNonLeftAlignedCells(current.lines, next.lines);
                List<TextLayoutEngine.TextLine> rhythmMiddle = current.lines;
                List<TextLayoutEngine.TextLine>? rhythmPrevious = previous;
                if (includedBridgeBandIndexes.Contains(current.idx - 1)) {
                    rhythmPrevious = bands[current.idx - 1];
                }
                if (next.idx == current.idx + 2) {
                    rhythmPrevious = current.lines;
                    rhythmMiddle = bands[current.idx + 1];
                }
                bool hasContinuousAlignedCells = HasEmphasizedText(bandSplits[start].lines[0]) &&
                                                 BandsHaveAlignedCells(current.lines, next.lines) &&
                                                 (BandsHaveCompatibleVerticalGap(current.lines, next.lines) ||
                                                  (rhythmPrevious is not null &&
                                                   BandsHaveCompatibleVerticalRhythm(
                                                       rhythmPrevious,
                                                       rhythmMiddle,
                                                       next.lines)));
                bool splitsAreSimilar = AreSplitsSimilar(baseSplits, next.splits);
                if (next.idx > current.idx + 2 ||
                    (!splitsAreSimilar &&
                     !hasNonLeftAlignedCells &&
                     !hasContinuousAlignedCells)) {
                    break;
                }
                InterveningBandDecision bridgeDecision = ClassifyInterveningBand(
                    bands,
                    current.idx,
                    next.idx,
                    baseSplits,
                    current.idx > bandSplits[start].idx
                        ? bands[current.idx - 1]
                        : null);
                if (bridgeDecision == InterveningBandDecision.Reject) {
                    break;
                }
                if (bridgeDecision == InterveningBandDecision.Include &&
                    IsUniformlyEmphasizedBand(bands[current.idx + 1]) &&
                    BandsAlignUsingSplits(bands[current.idx + 1], next.lines, next.splits)) {
                    break;
                }

                bool baseSplitsSeparateNextCells = SplitsSeparatePositionedCells(
                    next.lines,
                    baseSplits,
                    baseSplits.Count + 1);
                bool nextRequiresAlignedCellSplits = !baseSplitsSeparateNextCells ||
                                                     (!splitsAreSimilar &&
                                                      (hasNonLeftAlignedCells || hasContinuousAlignedCells));
                if (!alignedSplitAccumulator.AppendResult(
                    next.lines,
                    requiresAlignedCellSplits || nextRequiresAlignedCellSplits)) {
                    break;
                }
                requiresAlignedCellSplits |= nextRequiresAlignedCellSplits;

                if (bridgeDecision == InterveningBandDecision.Include) {
                    includedBridgeBandIndexes.Add(current.idx + 1);
                }
                end++;
            }
            // Build table for [start..end], including a compatible header-only band immediately above it.
            var groupLines = new List<TextLayoutEngine.TextLine>();
            if (headerLines is not null) {
                groupLines.AddRange(headerLines);
            }

            var splitBandIndexes = new HashSet<int>();
            for (int splitIndex = start; splitIndex <= end; splitIndex++) {
                splitBandIndexes.Add(bandSplits[splitIndex].idx);
            }
            for (int bandIndex = bandSplits[start].idx; bandIndex <= bandSplits[end].idx; bandIndex++) {
                if (splitBandIndexes.Contains(bandIndex) || includedBridgeBandIndexes.Contains(bandIndex)) {
                    groupLines.AddRange(bands[bandIndex]);
                }
            }
            List<double> effectiveSplits = baseSplits;
            if (requiresAlignedCellSplits) {
                List<double>? alignedSplits = alignedSplitAccumulator.GetSplits();
                if (alignedSplits is null) {
                    k = end + 1;
                    continue;
                }
                effectiveSplits = alignedSplits;
            }
            Dictionary<TextLayoutEngine.TextLine, List<double>>? lineSplitOverrides = null;
            if (headerLines is not null || includedBridgeBandIndexes.Count > 0) {
                lineSplitOverrides = new Dictionary<TextLayoutEngine.TextLine, List<double>>();
                if (headerLines is not null) {
                    foreach (TextLayoutEngine.TextLine headerLine in headerLines) {
                        lineSplitOverrides[headerLine] = baseSplits;
                    }
                }
                foreach (int bridgeBandIndex in includedBridgeBandIndexes) {
                    foreach (TextLayoutEngine.TextLine bridgeLine in bands[bridgeBandIndex]) {
                        lineSplitOverrides[bridgeLine] = baseSplits;
                    }
                }
            }
            var table = BuildTableFromLinesAndSplits(
                groupLines,
                effectiveSplits,
                "band-group",
                lineSplitOverrides);
            if (table != null &&
                (table.Rows.Count >= 3 || HasStrongTwoRowEvidence(table, groupLines)) &&
                HasValidatedRows(table, groupLines) &&
                !HasCrossingColumnText(groupLines, effectiveSplits, lineSplitOverrides)) {
                result.Add(table);
                claimedBandIndexes.UnionWith(splitBandIndexes);
                claimedBandIndexes.UnionWith(includedBridgeBandIndexes);
                if (headerLines is not null) claimedBandIndexes.Add(precedingBandIndex);
            }
            k = end + 1;
        }
        return result;
    }

    /// <summary>
    /// True when text assigned to one column runs past the start of text assigned to the next column on another row.
    /// Real columns are separated by a vertical channel; a chart title above its legend, set beside a side panel
    /// heading, only looks like a two-column header. Rows holding a single cell (a spanning title) are ignored.
    /// </summary>
    private static bool HasCrossingColumnText(
        List<TextLayoutEngine.TextLine> lines,
        List<double> splits,
        Dictionary<TextLayoutEngine.TextLine, List<double>>? lineSplitOverrides) {
        if (splits.Count == 0) return false;
        int columnCount = splits.Count + 1;
        var lefts = new double[columnCount];
        var rights = new double[columnCount];
        for (int index = 0; index < columnCount; index++) {
            lefts[index] = double.MaxValue;
            rights[index] = double.MinValue;
        }
        double fontSize = 0D;
        foreach (TextLayoutEngine.TextLine line in lines) {
            List<double> lineSplits = lineSplitOverrides is not null &&
                                      lineSplitOverrides.TryGetValue(line, out List<double>? splitOverride)
                ? splitOverride
                : splits;
            if (lineSplits.Count != splits.Count) continue;
            var lineLefts = new double[columnCount];
            var lineRights = new double[columnCount];
            var populated = new bool[columnCount];
            foreach (PdfTextSpan span in line.Spans) {
                if (string.IsNullOrWhiteSpace(span.Text)) continue;
                int column = 0;
                while (column < lineSplits.Count && span.X >= lineSplits[column]) column++;
                double right = span.X + Math.Max(0D, span.Advance);
                if (!populated[column]) {
                    lineLefts[column] = span.X;
                    lineRights[column] = right;
                    populated[column] = true;
                } else {
                    lineLefts[column] = Math.Min(lineLefts[column], span.X);
                    lineRights[column] = Math.Max(lineRights[column], right);
                }
                fontSize = Math.Max(fontSize, span.FontSize);
            }
            if (populated.Count(static value => value) < 2) continue;
            for (int column = 0; column < columnCount; column++) {
                if (!populated[column]) continue;
                lefts[column] = Math.Min(lefts[column], lineLefts[column]);
                rights[column] = Math.Max(rights[column], lineRights[column]);
            }
        }
        double tolerance = Math.Max(2D, fontSize * 0.5D);
        for (int column = 0; column + 1 < columnCount; column++) {
            if (rights[column] == double.MinValue || lefts[column + 1] == double.MaxValue) continue;
            if (rights[column] > lefts[column + 1] + tolerance) return true;
        }
        return false;
    }

    private enum InterveningBandDecision {
        Reject,
        Include,
        Skip
    }

    private static InterveningBandDecision ClassifyInterveningBand(
        List<List<TextLayoutEngine.TextLine>> bands,
        int currentBandIndex,
        int nextBandIndex,
        List<double> splits,
        List<TextLayoutEngine.TextLine>? establishedPreviousBand) {
        if (nextBandIndex == currentBandIndex + 1) return InterveningBandDecision.Skip;
        if (nextBandIndex != currentBandIndex + 2 || splits.Count == 0) {
            return InterveningBandDecision.Reject;
        }

        List<TextLayoutEngine.TextLine> intervening = bands[currentBandIndex + 1];
        if (intervening.Count != 1) return InterveningBandDecision.Reject;
        TextLayoutEngine.TextLine line = intervening[0];
        if (!HasCompatibleRowRhythm(
                establishedPreviousBand,
                bands[currentBandIndex],
                line,
                bands[nextBandIndex])) {
            return InterveningBandDecision.Reject;
        }
        if (!HasMeaningfulHorizontalOverlap(line, bands[currentBandIndex], bands[nextBandIndex])) {
            return InterveningBandDecision.Reject;
        }

        string[] cells = SplitBySplits(line, splits);
        // Producers emit a line as one show operator or one per word; classify by word
        // positions so a narrative line that wraps across a column split is not a single cell.
        if (cells.Count(static cell => !string.IsNullOrWhiteSpace(cell)) != 1 ||
            CountWordColumns(line, splits) > 1) {
            return InterveningBandDecision.Reject;
        }
        bool crossesColumnBoundary = CrossesColumnBoundary(line, splits);
        bool hasSectionLabelEvidence = HasEmphasizedText(line);
        bool hasCompactRowShape = LooksLikeCompactSpanningRow(line, cells);
        bool followsEmphasizedHeader = bands[currentBandIndex].Count == 1 &&
                                       HasEmphasizedText(bands[currentBandIndex][0]);
        if (crossesColumnBoundary) {
            return hasSectionLabelEvidence || hasCompactRowShape
                ? InterveningBandDecision.Include
                : InterveningBandDecision.Skip;
        }
        return followsEmphasizedHeader && (hasSectionLabelEvidence || hasCompactRowShape)
            ? InterveningBandDecision.Include
            : InterveningBandDecision.Reject;
    }

    private static int CountWordColumns(TextLayoutEngine.TextLine line, List<double> splits) {
        var columns = new HashSet<int>();
        foreach (PdfTextSpan span in line.Spans) {
            string text = span.Text ?? string.Empty;
            IReadOnlyList<double>? advances = span.CharacterAdvances;
            if (advances == null || advances.Count != text.Length || Math.Abs(span.RotationDegrees) > 0.5D) {
                columns.Add(ColumnAt(span.X, splits));
                continue;
            }
            double offset = 0D;
            bool inWord = false;
            for (int index = 0; index < text.Length; index++) {
                bool whitespace = char.IsWhiteSpace(text[index]);
                if (!whitespace && !inWord) columns.Add(ColumnAt(span.X + offset, splits));
                inWord = !whitespace;
                offset += advances[index];
            }
        }
        return columns.Count;
    }

    private static int ColumnAt(double x, List<double> splits) {
        int column = 0;
        while (column < splits.Count && x >= splits[column]) column++;
        return column;
    }

    private static bool LooksLikeCompactSpanningRow(
        TextLayoutEngine.TextLine line,
        IReadOnlyList<string> cells) {
        string value = ContentStructureExtractor.NormalizeShattered(
            cells.First(static cell => !string.IsNullOrWhiteSpace(cell))).Trim();
        return value.Length > 0 &&
               GetOccupiedWidthInFontSizes(line) <= MaximumCompactCellWidthInFontSizes &&
               !ContentStructureExtractor.EndsWithSentenceTerminal(value);
    }

    private static bool HasMeaningfulHorizontalOverlap(
        TextLayoutEngine.TextLine line,
        List<TextLayoutEngine.TextLine> previousBand,
        List<TextLayoutEngine.TextLine> nextBand) {
        double tableLeft = previousBand.Concat(nextBand).Min(static candidate => Math.Min(candidate.XStart, candidate.XEnd));
        double tableRight = previousBand.Concat(nextBand).Max(static candidate => Math.Max(candidate.XStart, candidate.XEnd));
        double lineLeft = Math.Min(line.XStart, line.XEnd);
        double lineRight = Math.Max(line.XStart, line.XEnd);
        double overlap = Math.Max(0D, Math.Min(lineRight, tableRight) - Math.Max(lineLeft, tableLeft));
        double narrowerWidth = Math.Min(lineRight - lineLeft, tableRight - tableLeft);
        return narrowerWidth > 0.001D && overlap + 0.001D >= narrowerWidth * 0.5D;
    }

    private static bool CrossesColumnBoundary(TextLayoutEngine.TextLine line, IReadOnlyList<double> splits) {
        double left = Math.Min(line.XStart, line.XEnd);
        double right = Math.Max(line.XStart, line.XEnd);
        return splits.Any(split => left < split - 1D && right > split + 1D);
    }

    private static bool BandsHaveAlignedCells(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand) {
        if (firstBand.Count != 1 || secondBand.Count != 1) return false;
        PositionedRow? first = TryCreatePositionedRow(firstBand[0]);
        PositionedRow? second = TryCreatePositionedRow(secondBand[0]);
        return first != null && second != null && PositionedRowsAlign(first, second);
    }

    private static bool BandsHaveCompatibleVerticalGap(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand,
        List<TextLayoutEngine.TextLine>? followingBand = null) {
        if (firstBand.Count != 1 || secondBand.Count == 0) return false;
        double secondBandTop = secondBand.Max(static line => line.Y);
        double gap = firstBand[0].Y - secondBandTop;
        if (gap <= 0D) return false;

        double largestFontSize = firstBand[0].Spans
            .Concat(secondBand.SelectMany(static line => line.Spans))
            .Select(static span => span.FontSize)
            .DefaultIfEmpty(0D)
            .Max();
        if (gap <= Math.Max(36D, largestFontSize * 3D)) return true;

        if (followingBand == null ||
            followingBand.Count == 0 ||
            !BandsContainAlignedCells(secondBand, followingBand)) {
            return false;
        }

        double followingGap = secondBandTop - followingBand.Max(static line => line.Y);
        if (followingGap <= 0D) return false;
        double smallerGap = Math.Min(gap, followingGap);
        double largerGap = Math.Max(gap, followingGap);
        return largerGap <= smallerGap * 1.75D;
    }

    private static bool BandsHaveNonLeftAlignedCells(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand) {
        if (firstBand.Count != 1 || secondBand.Count != 1) return false;
        PositionedRow? first = TryCreatePositionedRow(firstBand[0]);
        PositionedRow? second = TryCreatePositionedRow(secondBand[0]);
        if (first == null || second == null || first.Cells.Count != second.Cells.Count) return false;

        bool hasNonLeftAlignment = false;
        for (int index = 0; index < first.Cells.Count; index++) {
            PositionedCell expected = first.Cells[index];
            PositionedCell current = second.Cells[index];
            bool leftAligned = Math.Abs(expected.From - current.From) <= 16D;
            bool centerAligned = Math.Abs(
                (expected.From + expected.To) / 2D -
                (current.From + current.To) / 2D) <= 16D;
            bool rightAligned = Math.Abs(expected.To - current.To) <= 16D;
            if (!leftAligned && !centerAligned && !rightAligned) return false;
            if (!leftAligned && (centerAligned || rightAligned)) hasNonLeftAlignment = true;
        }
        return hasNonLeftAlignment;
    }

    private static bool HasCompatibleRowRhythm(
        List<TextLayoutEngine.TextLine>? establishedPreviousBand,
        List<TextLayoutEngine.TextLine> previousBand,
        TextLayoutEngine.TextLine intervening,
        List<TextLayoutEngine.TextLine> nextBand) {
        double previousY = previousBand.Average(static line => line.Y);
        double nextY = nextBand.Average(static line => line.Y);
        double upperGap = previousY - intervening.Y;
        double lowerGap = intervening.Y - nextY;
        if (upperGap <= 0D || lowerGap <= 0D) return false;
        double smaller = Math.Min(upperGap, lowerGap);
        double larger = Math.Max(upperGap, lowerGap);
        if (larger > smaller * 1.75D) return false;

        if (establishedPreviousBand is null || establishedPreviousBand.Count == 0) {
            return larger <= 36D;
        }

        double establishedPreviousY = establishedPreviousBand.Average(static candidate => candidate.Y);
        double establishedCurrentY = previousBand.Average(static candidate => candidate.Y);
        double establishedGap = establishedPreviousY - establishedCurrentY;
        return establishedGap > 0D && larger <= Math.Max(36D, establishedGap * 1.75D);
    }

    private static bool HasEmphasizedText(TextLayoutEngine.TextLine line) {
        PdfTextSpan[] spans = line.Spans
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
            .ToArray();
        return spans.Length > 0 && spans.All(static span => span.IsBold);
    }

    private static List<TextLayoutEngine.TextLine>? TryGetPrecedingHeaderLines(
        List<List<TextLayoutEngine.TextLine>> bands,
        int bodyBandIndex,
        List<double> bodySplits) {
        int headerBandIndex = bodyBandIndex - 1;
        if (headerBandIndex < 0 || bodySplits.Count == 0) {
            return null;
        }

        List<TextLayoutEngine.TextLine> headerBand = bands[headerBandIndex];
        if (headerBand.Count > 1) {
            // A header row whose labels are far apart (a wide label column next to right-aligned figures) is built
            // as several lines on one baseline. Treat them as the single row they are.
            TextLayoutEngine.TextLine? combined = TryCombineSameBaselineLines(headerBand);
            if (combined is null) return null;
            headerBand = new List<TextLayoutEngine.TextLine> { combined };
        }
        if (headerBand.Count != 1 || IsLeaderBand(headerBand)) {
            return null;
        }

        List<TextLayoutEngine.TextLine>? followingBodyBand = bodyBandIndex + 1 < bands.Count
            ? bands[bodyBandIndex + 1]
            : null;
        if ((!BandsHaveAlignedCells(headerBand, bands[bodyBandIndex]) &&
             !HasEmphasizedText(headerBand[0]))) {
            return null;
        }

        string[] headerCells = SplitBySplits(headerBand[0], bodySplits);
        if (!LooksLikeHeaderRow(headerCells)) {
            return null;
        }

        if (!BandsHaveCompatibleVerticalGap(headerBand, bands[bodyBandIndex], followingBodyBand)) {
            var twoRowLines = new List<TextLayoutEngine.TextLine>(headerBand.Count + bands[bodyBandIndex].Count);
            twoRowLines.AddRange(headerBand);
            twoRowLines.AddRange(bands[bodyBandIndex]);
            PdfStructuredTable? twoRowTable = BuildTableFromLinesAndSplits(
                twoRowLines,
                bodySplits,
                "band-group");
            if (twoRowTable is null || !HasStrongHeaderAndBodyEvidence(twoRowTable, twoRowLines)) {
                return null;
            }
        }

        return headerBand;
    }

    private static bool LooksLikeHeaderRow(string[] cells) {
        if (cells.Length < 2) {
            return false;
        }

        int labelCells = 0;
        for (int i = 0; i < cells.Length; i++) {
            string cell = ContentStructureExtractor.NormalizeShattered(cells[i]).Trim();
            if (cell.Length == 0) {
                return false;
            }
            if (PdfUnicodeScalarAnalysis.ContainsLetter(cell) || PdfUnicodeScalarAnalysis.IsAllDecimalDigits(cell)) {
                labelCells++;
                continue;
            }
            // Short symbol labels such as "#" or "%" head numbering and rate columns.
            if (!(cell.Length <= 2 && cell.All(static ch => char.IsSymbol(ch) || char.IsPunctuation(ch)))) {
                return false;
            }
        }

        return labelCells > 0;
    }

    private static TextLayoutEngine.TextLine? TryCombineSameBaselineLines(List<TextLayoutEngine.TextLine> lines) {
        double y = lines[0].Y;
        if (lines.Any(line => Math.Abs(line.Y - y) > 1.5D)) return null;
        TextLayoutEngine.TextLine[] ordered = lines.OrderBy(static line => line.XStart).ToArray();
        var spans = ordered.SelectMany(static line => line.Spans).ToList();
        if (spans.Count == 0) return null;
        return new TextLayoutEngine.TextLine(
            y,
            ordered[0].XStart,
            ordered.Max(static line => line.XEnd),
            string.Join(" ", ordered.Select(static line => line.Text)),
            spans,
            ordered[0].ReadingDirection);
    }

    private static bool IsUniformlyEmphasizedBand(List<TextLayoutEngine.TextLine> band) =>
        band.Count > 0 && band.All(static line => HasEmphasizedText(line));

    private static bool StartsStructurallySeparatedTable(
        List<TextLayoutEngine.TextLine> current,
        List<TextLayoutEngine.TextLine> candidateHeader,
        List<TextLayoutEngine.TextLine>? previous,
        List<TextLayoutEngine.TextLine>? following) {
        if (!IsUniformlyEmphasizedBand(candidateHeader) ||
            current.Count == 0 ||
            candidateHeader.Count == 0 ||
            following is null ||
            following.Count == 0 ||
            !BandsContainAlignedCells(candidateHeader, following)) {
            return false;
        }

        double boundaryGap = current.Average(static line => line.Y) - candidateHeader.Average(static line => line.Y);
        if (boundaryGap <= 0D) return false;
        double referenceGap = previous is null || previous.Count == 0
            ? candidateHeader.Average(static line => line.Y) - following.Average(static line => line.Y)
            : previous.Average(static line => line.Y) - current.Average(static line => line.Y);
        if (referenceGap <= 0D) return false;

        double fontSize = current
            .Concat(candidateHeader)
            .Concat(following)
            .SelectMany(static line => line.Spans)
            .Select(static span => span.FontSize)
            .DefaultIfEmpty(0D)
            .Max();
        return boundaryGap >= Math.Max(referenceGap * 1.35D, referenceGap + Math.Max(2D, fontSize * 0.35D));
    }

}
