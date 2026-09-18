namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private readonly struct TableRowTextSizing {
        internal TableRowTextSizing(double fontSize, double runFontSizeScale) {
            FontSize = fontSize;
            RunFontSizeScale = runFontSizeScale;
        }

        internal double FontSize { get; }

        internal double RunFontSizeScale { get; }
    }

    private sealed class TableCellTextWidthMeasurement {
        internal TableCellTextWidthMeasurement(
            TableTextLineWidthMeasurement[][] lines,
            double maxExplicitFontSize) {
            Lines = lines;
            MaxExplicitFontSize = maxExplicitFontSize;
        }

        internal TableTextLineWidthMeasurement[][] Lines { get; }

        internal double MaxExplicitFontSize { get; }
    }

    private readonly struct TableTextLineWidthMeasurement {
        internal TableTextLineWidthMeasurement(TableRunWidthMeasurement[] runs) {
            Runs = runs;
        }

        internal TableRunWidthMeasurement[] Runs { get; }
    }

    private readonly struct TableRunWidthMeasurement {
        internal TableRunWidthMeasurement(double horizontalOffset, double fixedWidth, double unitTextWidth, double? explicitFontSize, bool inline) {
            HorizontalOffset = horizontalOffset;
            FixedWidth = fixedWidth;
            UnitTextWidth = unitTextWidth;
            ExplicitFontSize = explicitFontSize;
            Inline = inline;
        }

        internal double HorizontalOffset { get; }

        internal double FixedWidth { get; }

        internal double UnitTextWidth { get; }

        internal double? ExplicitFontSize { get; }

        internal bool Inline { get; }
    }

    private static TableRowTextSizing ResolveTableRowTextSizing(TableBlock table, PdfTableStyle style, int rowIndex, int columnCount, double[] columnWidths, double columnGap, double rowFontSize, bool rowUsesBold, PdfOptions? options) {
        if (!style.ShrinkTextToFit || rowFontSize <= 0D) {
            return new TableRowTextSizing(rowFontSize, 1D);
        }

        double minimumFontSize = style.MinimumShrinkFontSize ?? 6D;
        PdfOptions effectiveOptions = options ?? new PdfOptions();
        PdfStandardFont rowFont = GetTableRowFont(effectiveOptions, rowUsesBold);
        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        TableCellTextWidthMeasurement?[]? explicitMeasurements = null;
        double resolvedFontSize = rowFontSize;
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            if (string.IsNullOrEmpty(cell.Text)) {
                continue;
            }

            if (minimumFontSize > rowFontSize) {
                continue;
            }

            double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
            double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column));
            double textWidth;
            if (GetMaxExplicitTableRunFontSize(cell) > 0D) {
                explicitMeasurements ??= new TableCellTextWidthMeasurement?[cells.Count];
                TableCellTextWidthMeasurement measurement = PrepareTableCellTextWidthMeasurement(cell, rowFont, effectiveOptions);
                explicitMeasurements[cellIndex] = measurement;
                textWidth = MeasurePreparedTableCellTextWidth(measurement, rowFontSize, 1D, minimumFontSize);
            } else {
                textWidth = MeasureTableCellTextWidth(cell, rowFont, rowFontSize, effectiveOptions);
            }
            if (textWidth <= innerWidth + 0.001D || textWidth <= 0.001D) {
                continue;
            }

            double candidate = Math.Max(minimumFontSize, rowFontSize * innerWidth / textWidth);
            resolvedFontSize = Math.Min(resolvedFontSize, candidate);
        }

        double scale = GetTableRunFontSizeScale(rowFontSize, resolvedFontSize);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            if (GetMaxExplicitTableRunFontSize(cell) <= resolvedFontSize + 0.001D) {
                continue;
            }

            TableCellTextWidthMeasurement measurement = explicitMeasurements?[cellIndex]
                ?? PrepareTableCellTextWidthMeasurement(cell, rowFont, effectiveOptions);
            double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
            double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column));
            double textWidth = MeasurePreparedTableCellTextWidth(measurement, resolvedFontSize, scale, minimumFontSize);
            if (textWidth <= innerWidth + 0.001D || textWidth <= 0.001D) {
                continue;
            }

            const double minimumScale = 0.001D;
            double minimumWidth = MeasurePreparedTableCellTextWidth(measurement, resolvedFontSize, minimumScale, minimumFontSize);
            if (minimumWidth > innerWidth + 0.001D) {
                scale = Math.Min(scale, minimumScale);
                continue;
            }

            double low = minimumScale;
            double high = scale;
            for (int iteration = 0; iteration < 20; iteration++) {
                double candidate = (low + high) / 2D;
                double candidateWidth = MeasurePreparedTableCellTextWidth(measurement, resolvedFontSize, candidate, minimumFontSize);
                if (candidateWidth <= innerWidth + 0.001D) {
                    low = candidate;
                } else {
                    high = candidate;
                }
            }

            scale = Math.Min(scale, low);
        }

        return new TableRowTextSizing(resolvedFontSize, scale);
    }

    private static double MeasureTableCellTextWidth(TableCellLayout cell, PdfStandardFont baseFont, double fontSize, PdfOptions options) {
        double width = 0D;
        if (cell.Paragraphs.Count > 0) {
            foreach (PdfTableCellParagraph paragraph in cell.Paragraphs) {
                width = Math.Max(width, MeasureTableRunsTextWidth(paragraph.Runs, baseFont, fontSize, options));
            }
        } else {
            width = MeasureTableRunsTextWidth(cell.Runs, baseFont, fontSize, options);
        }

        return width;
    }

    private static double MeasureTableRunsTextWidth(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions options) {
        System.Collections.Generic.IReadOnlyList<PdfTextRun> normalizedRuns = NormalizeFallbackRuns(runs, baseFont, options);
        double width = 0D;
        foreach (System.Collections.Generic.IReadOnlyList<PdfTextRun> line in BuildPageTextLineRuns(normalizedRuns)) {
            width = Math.Max(width, MeasurePageTextLineRuns(line, baseFont, fontSize, options));
        }

        return width;
    }

    private static TableCellTextWidthMeasurement PrepareTableCellTextWidthMeasurement(TableCellLayout cell, PdfStandardFont baseFont, PdfOptions options) {
        int paragraphCount = cell.Paragraphs.Count;
        var lines = new TableTextLineWidthMeasurement[Math.Max(1, paragraphCount)][];
        double maxExplicitFontSize = 0D;
        if (paragraphCount > 0) {
            for (int paragraphIndex = 0; paragraphIndex < paragraphCount; paragraphIndex++) {
                System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = cell.Paragraphs[paragraphIndex].Runs;
                lines[paragraphIndex] = PrepareTableTextLineWidths(runs, baseFont, options);
                maxExplicitFontSize = Math.Max(maxExplicitFontSize, GetMaxExplicitRunFontSize(runs));
            }
        } else {
            lines[0] = PrepareTableTextLineWidths(cell.Runs, baseFont, options);
            maxExplicitFontSize = GetMaxExplicitRunFontSize(cell.Runs);
        }

        return new TableCellTextWidthMeasurement(lines, maxExplicitFontSize);
    }

    private static TableTextLineWidthMeasurement[] PrepareTableTextLineWidths(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, PdfOptions options) {
        System.Collections.Generic.IReadOnlyList<PdfTextRun> normalizedRuns = NormalizeFallbackRuns(runs, baseFont, options);
        System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTextRun>> sourceLines = BuildPageTextLineRuns(normalizedRuns);
        var preparedLines = new TableTextLineWidthMeasurement[sourceLines.Count];
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            System.Collections.Generic.IReadOnlyList<PdfTextRun> sourceRuns = sourceLines[lineIndex];
            var preparedRuns = new TableRunWidthMeasurement[sourceRuns.Count];
            for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                PdfTextRun run = sourceRuns[runIndex];
                if (run.InlineElement != null) {
                    preparedRuns[runIndex] = new TableRunWidthMeasurement(
                        run.HorizontalOffset,
                        run.InlineElement.Width,
                        0D,
                        null,
                        inline: true);
                    continue;
                }

                PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                    ? resolvedNamedFont
                    : null;
                double unitTextWidth = MeasureRichText(
                    run.Text ?? string.Empty,
                    ResolvePageTextRunFont(run, baseFont),
                    namedFont,
                    1D,
                    run.Baseline,
                    options);
                preparedRuns[runIndex] = new TableRunWidthMeasurement(
                    run.HorizontalOffset,
                    0D,
                    unitTextWidth,
                    run.FontSize,
                    inline: false);
            }

            preparedLines[lineIndex] = new TableTextLineWidthMeasurement(preparedRuns);
        }

        return preparedLines;
    }

    private static double MeasurePreparedTableCellTextWidth(TableCellTextWidthMeasurement measurement, double fontSize, double runFontSizeScale, double minimumShrinkFontSize) {
        double width = 0D;
        for (int paragraphIndex = 0; paragraphIndex < measurement.Lines.Length; paragraphIndex++) {
            TableTextLineWidthMeasurement[] lines = measurement.Lines[paragraphIndex];
            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
                width = Math.Max(width, MeasurePreparedTableTextLineWidth(lines[lineIndex], fontSize, runFontSizeScale, minimumShrinkFontSize));
            }
        }

        return width;
    }

    private static double MeasurePreparedTableTextLineWidth(TableTextLineWidthMeasurement line, double fontSize, double runFontSizeScale, double minimumShrinkFontSize) {
        double width = 0D;
        double minimumExplicitFontSize = minimumShrinkFontSize > 0D ? minimumShrinkFontSize : 0.001D;
        for (int runIndex = 0; runIndex < line.Runs.Length; runIndex++) {
            TableRunWidthMeasurement run = line.Runs[runIndex];
            if (run.Inline) {
                width += run.HorizontalOffset + run.FixedWidth;
                continue;
            }

            if (runFontSizeScale >= 0.999D) {
                width += run.HorizontalOffset;
            }
            double effectiveFontSize = run.ExplicitFontSize.HasValue
                ? run.ExplicitFontSize.Value <= minimumExplicitFontSize
                    ? run.ExplicitFontSize.Value
                    : Math.Max(minimumExplicitFontSize, run.ExplicitFontSize.Value * runFontSizeScale)
                : fontSize;
            width += run.UnitTextWidth * effectiveFontSize;
        }

        return width;
    }

    private static double GetTableRunFontSizeScale(double originalFontSize, double resolvedFontSize) {
        if (originalFontSize <= 0D ||
            resolvedFontSize >= originalFontSize - 0.001D) {
            return 1D;
        }

        return resolvedFontSize / originalFontSize;
    }

    private static double GetMaxExplicitTableRunFontSize(TableCellLayout cell) {
        double max = GetMaxExplicitRunFontSize(cell.Runs);
        for (int paragraphIndex = 0; paragraphIndex < cell.Paragraphs.Count; paragraphIndex++) {
            max = Math.Max(max, GetMaxExplicitRunFontSize(cell.Paragraphs[paragraphIndex].Runs));
        }

        return max;
    }

    private static double GetMaxExplicitRunFontSize(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs) {
        double max = 0D;
        foreach (PdfTextRun run in runs) {
            if (run.FontSize.HasValue) {
                max = Math.Max(max, run.FontSize.Value);
            }
        }

        return max;
    }
}
