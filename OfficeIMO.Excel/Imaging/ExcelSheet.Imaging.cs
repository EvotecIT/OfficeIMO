using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates a format-neutral visual snapshot for a worksheet range or the used range.
        /// </summary>
        public ExcelRangeVisualSnapshot CreateVisualSnapshot(ExcelWorksheetImageExportOptions? options = null) {
            ExcelWorksheetImageExportOptions resolved = NormalizeWorksheetOptions(options);
            WorksheetImageRangeResolution range = ResolveWorksheetImageRanges(resolved, allowMultipleResults: false)[0];
            return ExcelRangeVisualSnapshotBuilder.Build(this, range.Range, resolved, range.Diagnostics);
        }

        /// <summary>
        /// Exports a worksheet range or used range as a supported raster format or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, ExcelWorksheetImageExportOptions? options = null) {
            ExcelWorksheetImageExportOptions resolved = NormalizeWorksheetOptions(options);
            WorksheetImageRangeResolution range = ResolveWorksheetImageRanges(resolved, allowMultipleResults: false)[0];
            ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(this, range.Range, resolved, range.Diagnostics);
            return ExcelRangeImageRenderer.Render(snapshot, format, resolved);
        }

        /// <summary>
        /// Exports one or more worksheet image results. Multi-area print areas and manual page-break splits are returned as separate images when requested.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, ExcelWorksheetImageExportOptions? options = null) {
            ExcelWorksheetImageExportOptions resolved = NormalizeWorksheetOptions(options);
            IReadOnlyList<WorksheetImageRangeResolution> ranges = ResolveWorksheetImageRanges(resolved, allowMultipleResults: true);
            var results = new List<OfficeImageExportResult>(ranges.Count);
            for (int index = 0; index < ranges.Count; index++) {
                results.Add(RenderWorksheetImageResult(format, ranges[index], resolved, index + 1, ranges.Count));
            }

            return results.AsReadOnly();
        }

        /// <summary>
        /// Renders the worksheet used range to PNG bytes.
        /// </summary>
        public byte[] ToPng(ExcelWorksheetImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders the worksheet used range to SVG text.
        /// </summary>
        public string ToSvg(ExcelWorksheetImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves the worksheet used range as PNG.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(string path, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsPng().Save(path);

        /// <summary>
        /// Saves the worksheet used range as SVG.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(string path, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsSvg().Save(path);

        /// <summary>
        /// Writes the worksheet used range as PNG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsPng().Save(stream);

        /// <summary>
        /// Writes the worksheet used range as SVG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsSvg().Save(stream);

        /// <summary>Asynchronously saves the worksheet used range as PNG.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            string path,
            ExcelWorksheetImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsPng().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the worksheet used range as SVG.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            string path,
            ExcelWorksheetImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsSvg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes the worksheet used range as PNG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            Stream stream,
            ExcelWorksheetImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsPng().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the worksheet used range as SVG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            Stream stream,
            ExcelWorksheetImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelWorksheetImageExportBuilder(this, options).AsSvg().SaveAsync(stream, cancellationToken);

        private static ExcelWorksheetImageExportOptions NormalizeWorksheetOptions(ExcelWorksheetImageExportOptions? options) {
            ExcelWorksheetImageExportOptions resolved = options?.CloneWorksheet() ?? new ExcelWorksheetImageExportOptions();
            resolved.ConditionalFormattingDate ??= DateTime.Today;
            resolved.HeaderFooterDateTime ??= DateTime.Now;
            resolved.Validate();

            return resolved;
        }

        private IReadOnlyList<WorksheetImageRangeResolution> ResolveWorksheetImageRanges(ExcelWorksheetImageExportOptions options, bool allowMultipleResults) {
            if (!string.IsNullOrWhiteSpace(options.Range)) {
                if (TryNormalizeWorksheetImageRange(options.Range!, out string? normalizedRange)) {
                    return ApplyManualPageBreakSplits(
                        SingleImageRange(normalizedRange!, Array.Empty<OfficeImageExportDiagnostic>()),
                        options,
                        allowMultipleResults);
                }

                throw new ArgumentException("Worksheet image export range must be a supported A1 cell or range reference.", nameof(options));
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>();
            if (options.UsePrintArea) {
                string? printArea = GetPrintArea();
                string source = Name + "!_xlnm.Print_Area";
                if (string.IsNullOrWhiteSpace(printArea)) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(
                        OfficeImageExportDiagnosticSeverity.Info,
                        ExcelImageExportDiagnosticCodes.PrintAreaMissing,
                        "Worksheet image export requested the print area, but no worksheet print area is configured; exporting the worksheet used range instead.",
                        source));
                } else {
                    List<string> printAreaParts = SplitDefinedNameParts(printArea!).ToList();
                    if (printAreaParts.Count > 1) {
                        if (allowMultipleResults && TryNormalizeWorksheetImageRanges(printAreaParts, out IReadOnlyList<string>? normalizedRanges)) {
                            return ApplyManualPageBreakSplits(
                                normalizedRanges!
                                .Select(range => new WorksheetImageRangeResolution(
                                    range,
                                    new[] {
                                        new OfficeImageExportDiagnostic(
                                            OfficeImageExportDiagnosticSeverity.Info,
                                            ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasSplit,
                                            "Multi-area worksheet print area was exported as separate image results.",
                                            source)
                                    }))
                                .ToList()
                                .AsReadOnly(),
                                options,
                                allowMultipleResults);
                        }

                        diagnostics.Add(new OfficeImageExportDiagnostic(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasUnsupported,
                            "Multi-area worksheet print areas are not supported by single-image export; exporting the worksheet used range instead.",
                            source));
                    } else if (TryNormalizeWorksheetImageRange(printArea!, out string? normalizedPrintArea)) {
                        return ApplyManualPageBreakSplits(SingleImageRange(normalizedPrintArea!, diagnostics), options, allowMultipleResults);
                    } else {
                        diagnostics.Add(new OfficeImageExportDiagnostic(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.PrintAreaUnsupported,
                            "Worksheet print area could not be parsed as a supported A1 range; exporting the worksheet used range instead.",
                            source));
                    }
                }
            }

            return ApplyManualPageBreakSplits(SingleImageRange(ResolveWorksheetUsedImageRange(options), diagnostics), options, allowMultipleResults);
        }

        private static IReadOnlyList<WorksheetImageRangeResolution> SingleImageRange(string range, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) =>
            new[] { new WorksheetImageRangeResolution(range, diagnostics) };

        private IReadOnlyList<WorksheetImageRangeResolution> ApplyManualPageBreakSplits(
            IReadOnlyList<WorksheetImageRangeResolution> ranges,
            ExcelWorksheetImageExportOptions options,
            bool allowMultipleResults) {
            if (!options.SplitByManualPageBreaks) {
                return ranges;
            }

            IReadOnlyList<OfficeImageExportDiagnostic> pageDiagnostics = BuildPageLevelUnsupportedDiagnostics(
                includePrintTitlesUnsupported: !allowMultipleResults,
                includeHeaderFooterUnsupported: !allowMultipleResults || !CanRenderHeaderFooterTextChrome(options.HeaderFooterDateTime ?? DateTime.Now));
            if (!allowMultipleResults) {
                return ranges
                    .Select(range => range
                        .WithDiagnostics(pageDiagnostics)
                        .WithDiagnostic(new OfficeImageExportDiagnostic(
                            OfficeImageExportDiagnosticSeverity.Warning,
                            ExcelImageExportDiagnosticCodes.ManualPageBreaksSingleImageUnsupported,
                            "Manual worksheet page-break splitting was requested through a single-image export path; exporting one image for the resolved range instead.",
                            Name + "!" + range.Range)))
                    .ToList()
                    .AsReadOnly();
            }

            var splitRanges = new List<WorksheetImageRangeResolution>();
            foreach (WorksheetImageRangeResolution range in ranges) {
                IReadOnlyList<string> pages = SplitRangeByManualPageBreaks(range.Range);
                if (pages.Count <= 1) {
                    splitRanges.Add(range.WithDiagnostics(pageDiagnostics));
                    continue;
                }

                foreach (string pageRange in pages) {
                    splitRanges.Add(range
                        .WithDiagnostics(pageDiagnostics)
                        .WithRangeAndDiagnostic(
                            pageRange,
                            new OfficeImageExportDiagnostic(
                                OfficeImageExportDiagnosticSeverity.Info,
                                ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit,
                                "Manual worksheet page breaks were used to split the image export into separate results.",
                                Name + "!" + range.Range)));
                }
            }

            return splitRanges.AsReadOnly();
        }

        private IReadOnlyList<OfficeImageExportDiagnostic> BuildPageLevelUnsupportedDiagnostics(bool includePrintTitlesUnsupported, bool includeHeaderFooterUnsupported) {
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            ExcelPrintTitles printTitles = GetPrintTitles();
            if (includePrintTitlesUnsupported && (printTitles.HasRows || printTitles.HasColumns)) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported,
                    "Worksheet print title rows or columns are configured, but image page output does not repeat them yet.",
                    Name + "!_xlnm.Print_Titles"));
            }

            ExcelSheetPageSetup pageSetup = GetPageSetup();
            if (ExcelPageSetupGeometry.HasUnsupportedFitToPageScale(pageSetup)) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.PageSetupUnsupported,
                    "Worksheet fit-to-width or fit-to-height page setup requests more than one page in a dimension, but image page output does not calculate automatic multi-page fit pagination yet.",
                    Name + "!pageSetup"));
            }

            if (includeHeaderFooterUnsupported && HasHeaderFooterContent()) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported,
                    "Worksheet headers or footers are configured, but image page output does not render page header/footer chrome yet.",
                    Name + "!headerFooter"));
            }

            return diagnostics.AsReadOnly();
        }

        private bool HasHeaderFooterContent() {
            HeaderFooter? headerFooter = WorksheetRoot.GetFirstChild<HeaderFooter>();
            if (headerFooter == null) {
                return false;
            }

            return HasText(headerFooter.OddHeader?.Text) ||
                HasText(headerFooter.OddFooter?.Text) ||
                HasText(headerFooter.EvenHeader?.Text) ||
                HasText(headerFooter.EvenFooter?.Text) ||
                HasText(headerFooter.FirstHeader?.Text) ||
                HasText(headerFooter.FirstFooter?.Text);
        }

        private static bool HasText(string? text) => !string.IsNullOrWhiteSpace(text);

        private IReadOnlyList<string> SplitRangeByManualPageBreaks(string range) {
            if (!A1.TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return new[] { range };
            }

            IReadOnlyList<PageBreakSegment> rowSegments = BuildPageBreakSegments(firstRow, lastRow, GetManualRowPageBreaks());
            IReadOnlyList<PageBreakSegment> columnSegments = BuildPageBreakSegments(firstColumn, lastColumn, GetManualColumnPageBreaks());
            if (rowSegments.Count == 1 && columnSegments.Count == 1) {
                return new[] { range };
            }

            var ranges = new List<string>(rowSegments.Count * columnSegments.Count);
            ExcelPageOrder pageOrder = GetPageSetup().PageOrder ?? ExcelPageOrder.DownThenOver;
            if (pageOrder == ExcelPageOrder.OverThenDown) {
                foreach (PageBreakSegment row in rowSegments) {
                    foreach (PageBreakSegment column in columnSegments) {
                        ranges.Add(ToRange(row.Start, column.Start, row.End, column.End));
                    }
                }
            } else {
                foreach (PageBreakSegment column in columnSegments) {
                    foreach (PageBreakSegment row in rowSegments) {
                        ranges.Add(ToRange(row.Start, column.Start, row.End, column.End));
                    }
                }
            }

            return ranges.AsReadOnly();
        }

        private static IReadOnlyList<PageBreakSegment> BuildPageBreakSegments(int first, int last, IReadOnlyList<int> breakAfterIndexes) {
            var segments = new List<PageBreakSegment>();
            int start = first;
            foreach (int breakAfter in breakAfterIndexes.Where(value => value >= first && value < last).Distinct().OrderBy(value => value)) {
                segments.Add(new PageBreakSegment(start, breakAfter));
                start = breakAfter + 1;
            }

            segments.Add(new PageBreakSegment(start, last));
            return segments.AsReadOnly();
        }

        private static string ToRange(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
            A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);

        private static bool TryNormalizeWorksheetImageRanges(IEnumerable<string> ranges, out IReadOnlyList<string>? normalizedRanges) {
            var normalized = new List<string>();
            foreach (string range in ranges) {
                if (!TryNormalizeWorksheetImageRange(range, out string? normalizedRange)) {
                    normalizedRanges = null;
                    return false;
                }

                normalized.Add(normalizedRange!);
            }

            normalizedRanges = normalized.AsReadOnly();
            return normalized.Count > 0;
        }

        private string ResolveWorksheetUsedImageRange(ExcelWorksheetImageExportOptions options) {
            string range = GetUsedRangeA1();
            if (!A1.TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return range;
            }

            IReadOnlyList<ExcelColumnSnapshot> columns = GetColumnDefinitions();
            Dictionary<int, ExcelRowSnapshot> rows = GetRowDefinitions().ToDictionary(row => row.Index);
            bool defaultRowsHidden = DefaultRowsHidden;
            if (options.IncludeImages) {
                foreach (ExcelImage image in Images) {
                    if (image.TryGetAbsoluteAnchorBounds(out int absoluteX, out int absoluteY, out int absoluteWidth, out int absoluteHeight)) {
                        ExpandAbsoluteVisualAnchor(absoluteX, absoluteY, absoluteWidth, absoluteHeight, columns, rows, options, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);
                    } else if (options.IncludeHidden || !IsHiddenAnchor(image.RowIndex, image.ColumnIndex, rows, defaultRowsHidden, columns)) {
                        ExpandVisualAnchor(
                            image.RowIndex,
                            image.ColumnIndex,
                            Math.Max(1, image.WidthPixels + Math.Max(0, image.OffsetXPixels)),
                            Math.Max(1, image.HeightPixels + Math.Max(0, image.OffsetYPixels)),
                            columns,
                            rows,
                            defaultRowsHidden,
                            options,
                            ref firstRow,
                            ref firstColumn,
                            ref lastRow,
                            ref lastColumn);
                    }
                }
            }

            if (options.IncludeCharts) {
                foreach (ExcelChart chart in Charts) {
                    if (chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                        if (chart.TryGetAbsoluteAnchorBounds(out int absoluteX, out int absoluteY, out int absoluteWidth, out int absoluteHeight)) {
                            ExpandAbsoluteVisualAnchor(absoluteX, absoluteY, absoluteWidth, absoluteHeight, columns, rows, options, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);
                        } else if (options.IncludeHidden || !IsHiddenAnchor(snapshot.RowIndex, snapshot.ColumnIndex, rows, defaultRowsHidden, columns)) {
                            ExpandVisualAnchor(
                                snapshot.RowIndex,
                                snapshot.ColumnIndex,
                                Math.Max(1, snapshot.WidthPixels + Math.Max(0, snapshot.OffsetXPixels)),
                                Math.Max(1, snapshot.HeightPixels + Math.Max(0, snapshot.OffsetYPixels)),
                                columns,
                                rows,
                                defaultRowsHidden,
                                options,
                                ref firstRow,
                                ref firstColumn,
                                ref lastRow,
                                ref lastColumn);
                        }
                    }
                }
            }

            if (options.IncludeDrawingObjects) {
                foreach (ExcelWorksheetDrawingObjectInfo drawing in ExcelWorksheetDrawingObjectResolver.FindDrawingObjects(WorksheetPart)) {
                    if (drawing.IsRenderable && (options.IncludeHidden || !IsHiddenAnchor(drawing.Row, drawing.Column, rows, defaultRowsHidden, columns))) {
                        ExpandVisualAnchor(
                            drawing.Row,
                            drawing.Column,
                            Math.Max(1, drawing.WidthPixels + Math.Max(0, drawing.OffsetXPixels)),
                            Math.Max(1, drawing.HeightPixels + Math.Max(0, drawing.OffsetYPixels)),
                            columns,
                            rows,
                            defaultRowsHidden,
                            options,
                            ref firstRow,
                            ref firstColumn,
                            ref lastRow,
                            ref lastColumn);
                    }
                }
            }

            return A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
        }

        private static bool TryNormalizeWorksheetImageRange(string range, out string? normalizedRange) {
            string withoutSheet = StripSheetPrefix(range).Replace("$", string.Empty).Trim();
            if (A1.TryParseRange(withoutSheet, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                normalizedRange = A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
                return true;
            }

            (int Row, int Col) cell = A1.ParseCellRef(withoutSheet);
            if (cell.Row > 0 && cell.Col > 0) {
                string reference = A1.CellReference(cell.Row, cell.Col);
                normalizedRange = reference + ":" + reference;
                return true;
            }

            normalizedRange = null;
            return false;
        }

        private static void ExpandVisualAnchor(
            int rowIndex,
            int columnIndex,
            int widthPixels,
            int heightPixels,
            IReadOnlyList<ExcelColumnSnapshot> columns,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rows,
            bool defaultRowsHidden,
            ExcelImageExportOptions options,
            ref int firstRow,
            ref int firstColumn,
            ref int lastRow,
            ref int lastColumn) {
            if (rowIndex <= 0 || columnIndex <= 0) {
                return;
            }

            firstRow = Math.Min(firstRow, rowIndex);
            firstColumn = Math.Min(firstColumn, columnIndex);
            lastRow = Math.Max(lastRow, ResolveLastVisualRow(rowIndex, heightPixels, rows, defaultRowsHidden, options));
            lastColumn = Math.Max(lastColumn, ResolveLastVisualColumn(columnIndex, widthPixels, columns, options));
        }

        private static void ExpandAbsoluteVisualAnchor(
            int xPixels,
            int yPixels,
            int widthPixels,
            int heightPixels,
            IReadOnlyList<ExcelColumnSnapshot> columns,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rows,
            ExcelImageExportOptions options,
            ref int firstRow,
            ref int firstColumn,
            ref int lastRow,
            ref int lastColumn) {
            int startColumn = ResolveColumnAtAbsoluteOffset(xPixels, columns, options);
            int endColumn = ResolveColumnAtAbsoluteOffset(xPixels + Math.Max(1, widthPixels), columns, options);
            int startRow = ResolveRowAtAbsoluteOffset(yPixels, rows, options);
            int endRow = ResolveRowAtAbsoluteOffset(yPixels + Math.Max(1, heightPixels), rows, options);
            firstColumn = Math.Min(firstColumn, startColumn);
            lastColumn = Math.Max(lastColumn, endColumn);
            firstRow = Math.Min(firstRow, startRow);
            lastRow = Math.Max(lastRow, endRow);
        }

        private static int ResolveColumnAtAbsoluteOffset(int offsetPixels, IReadOnlyList<ExcelColumnSnapshot> columns, ExcelImageExportOptions options) {
            double cursor = 0D;
            for (int column = 1; column < 16384; column++) {
                ExcelColumnSnapshot? definition = columns.FirstOrDefault(item => column >= item.StartIndex && column <= item.EndIndex);
                double width = ResolveColumnWidth(definition, options);
                if (definition?.Hidden == true && !options.IncludeHidden) {
                    width = 0D;
                }

                if (cursor + width >= offsetPixels) {
                    return column;
                }

                cursor += width;
            }

            return 16384;
        }

        private static int ResolveRowAtAbsoluteOffset(int offsetPixels, IReadOnlyDictionary<int, ExcelRowSnapshot> rows, ExcelImageExportOptions options) {
            double cursor = 0D;
            for (int row = 1; row < 1048576; row++) {
                rows.TryGetValue(row, out ExcelRowSnapshot? definition);
                double height = ResolveRowHeight(definition, options);
                if (definition?.Hidden == true && !options.IncludeHidden) {
                    height = 0D;
                }

                if (cursor + height >= offsetPixels) {
                    return row;
                }

                cursor += height;
            }

            return 1048576;
        }

        private static int ResolveLastVisualColumn(int startColumn, int widthPixels, IReadOnlyList<ExcelColumnSnapshot> columns, ExcelImageExportOptions options) {
            double remaining = Math.Max(1D, widthPixels);
            int column = startColumn;
            while (column < 16384) {
                remaining -= ResolveVisibleColumnWidth(column, columns, options);
                if (remaining <= 0D) {
                    return column;
                }

                column++;
            }

            return column;
        }

        private static int ResolveLastVisualRow(int startRow, int heightPixels, IReadOnlyDictionary<int, ExcelRowSnapshot> rows, bool defaultRowsHidden, ExcelImageExportOptions options) {
            double remaining = Math.Max(1D, heightPixels);
            int row = startRow;
            while (row < 1048576) {
                rows.TryGetValue(row, out ExcelRowSnapshot? definition);
                remaining -= ResolveVisibleRowHeight(row, definition, rows, defaultRowsHidden, options);
                if (remaining <= 0D) {
                    return row;
                }

                row++;
            }

            return row;
        }

        private static bool IsHiddenAnchor(int rowIndex, int columnIndex, IReadOnlyDictionary<int, ExcelRowSnapshot> rows, bool defaultRowsHidden, IReadOnlyList<ExcelColumnSnapshot> columns) =>
            IsHiddenRow(rowIndex, rows, defaultRowsHidden) || IsHiddenColumn(columnIndex, columns);

        private static bool IsHiddenRow(int rowIndex, IReadOnlyDictionary<int, ExcelRowSnapshot> rows, bool defaultRowsHidden) =>
            rows.TryGetValue(rowIndex, out ExcelRowSnapshot? definition) ? definition.Hidden : defaultRowsHidden;

        private static bool IsHiddenColumn(int columnIndex, IReadOnlyList<ExcelColumnSnapshot> columns) =>
            columns.Any(definition => definition.Hidden && columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex);

        private static double ResolveVisibleColumnWidth(int column, IReadOnlyList<ExcelColumnSnapshot> columns, ExcelImageExportOptions options) {
            ExcelColumnSnapshot? definition = columns.FirstOrDefault(item => column >= item.StartIndex && column <= item.EndIndex);
            if (definition?.Hidden == true && !options.IncludeHidden) {
                return 0D;
            }

            return ResolveColumnWidth(definition, options);
        }

        private static double ResolveVisibleRowHeight(int rowIndex, ExcelRowSnapshot? definition, IReadOnlyDictionary<int, ExcelRowSnapshot> rows, bool defaultRowsHidden, ExcelImageExportOptions options) {
            if (IsHiddenRow(rowIndex, rows, defaultRowsHidden) && !options.IncludeHidden) {
                return 0D;
            }

            return ResolveRowHeight(definition, options);
        }

        private static double ResolveColumnWidth(ExcelColumnSnapshot? definition, ExcelImageExportOptions options) {
            if (definition?.Width == null) {
                return options.DefaultColumnWidthPixels;
            }

            return Math.Max(1D, Math.Round((definition.Width.Value * 7D) + 5D, 2));
        }

        private static double ResolveRowHeight(ExcelRowSnapshot? definition, ExcelImageExportOptions options) {
            if (definition?.Height == null) {
                return options.DefaultRowHeightPixels;
            }

            return Math.Max(1D, Math.Round(definition.Height.Value * 96D / 72D, 2));
        }

        private sealed class WorksheetImageRangeResolution {
            internal WorksheetImageRangeResolution(string range, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
                Range = range;
                Diagnostics = diagnostics;
            }

            internal string Range { get; }
            internal IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

            internal WorksheetImageRangeResolution WithDiagnostic(OfficeImageExportDiagnostic diagnostic) =>
                WithRangeAndDiagnostic(Range, diagnostic);

            internal WorksheetImageRangeResolution WithDiagnostics(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
                if (diagnostics.Count == 0) {
                    return this;
                }

                var combined = new List<OfficeImageExportDiagnostic>(Diagnostics.Count + diagnostics.Count);
                combined.AddRange(Diagnostics);
                combined.AddRange(diagnostics);
                return new WorksheetImageRangeResolution(Range, combined.AsReadOnly());
            }

            internal WorksheetImageRangeResolution WithRangeAndDiagnostic(string range, OfficeImageExportDiagnostic diagnostic) {
                var diagnostics = new List<OfficeImageExportDiagnostic>(Diagnostics.Count + 1);
                diagnostics.AddRange(Diagnostics);
                diagnostics.Add(diagnostic);
                return new WorksheetImageRangeResolution(range, diagnostics.AsReadOnly());
            }
        }

        private readonly struct PageBreakSegment {
            internal PageBreakSegment(int start, int end) {
                Start = start;
                End = end;
            }

            internal int Start { get; }
            internal int End { get; }
        }
    }
}
