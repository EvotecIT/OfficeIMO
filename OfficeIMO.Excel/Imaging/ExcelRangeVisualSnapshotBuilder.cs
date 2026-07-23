using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    internal sealed class ExcelSourceImageBudget {
        internal ExcelSourceImageBudget(long maximumBytes) {
            if (maximumBytes < 0L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            RemainingBytes = maximumBytes;
        }

        internal long RemainingBytes { get; private set; }

        internal void Consume(long byteCount) {
            if (byteCount < 0L || byteCount > RemainingBytes) {
                throw new InvalidOperationException("The aggregate source-image budget was exceeded.");
            }
            RemainingBytes -= byteCount;
        }
    }

    internal static class ExcelRangeVisualSnapshotBuilder {
        private const int MaxSparklineDataCells = 100_000;
        internal static ExcelRangeVisualSnapshot Build(ExcelSheet sheet, string range, ExcelImageExportOptions options, IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null, ExcelSourceImageBudget? sourceImageBudget = null) {
            if (sheet == null) {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            if (!A1.TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                throw new ArgumentException("Range must be a valid A1 range.", nameof(range));
            }

            if (options.MaximumRenderedCells <= 0) {
                throw new ArgumentOutOfRangeException(
                    nameof(options),
                    options.MaximumRenderedCells,
                    "MaximumRenderedCells must be greater than zero.");
            }
            sourceImageBudget ??= new ExcelSourceImageBudget(options.MaximumTotalSourceImageBytes);

            long rowCount = (long)lastRow - firstRow + 1L;
            long columnCount = (long)lastColumn - firstColumn + 1L;
            if (rowCount <= 0L || columnCount <= 0L || rowCount > options.MaximumRenderedCells / columnCount) {
                throw new InvalidOperationException(
                    "The requested Excel visual range exceeds the configured limit of " +
                    options.MaximumRenderedCells.ToString(CultureInfo.InvariantCulture) +
                    " rendered cells.");
            }

            var diagnostics = initialDiagnostics == null
                ? new List<OfficeImageExportDiagnostic>()
                : new List<OfficeImageExportDiagnostic>(initialDiagnostics);
            Dictionary<int, ExcelRowSnapshot> rowDefinitions = sheet.GetRowDefinitions().ToDictionary(row => row.Index);
            bool defaultRowsHidden = sheet.DefaultRowsHidden;
            List<ExcelColumnSnapshot> columnDefinitions = sheet.GetColumnDefinitions().ToList();
            List<ExcelMergedRangeSnapshot> merges = sheet.GetMergedRanges()
                .Where(merge => Intersects(merge.StartRow, merge.StartColumn, merge.EndRow, merge.EndColumn, firstRow, firstColumn, lastRow, lastColumn))
                .ToList();

            var columns = new List<ExcelVisualColumn>();
            double x = 0D;
            for (int column = firstColumn; column <= lastColumn; column++) {
                ExcelColumnSnapshot? definition = columnDefinitions.FirstOrDefault(item => column >= item.StartIndex && column <= item.EndIndex);
                if (definition?.Hidden == true && !options.IncludeHidden) {
                    continue;
                }

                double width = ResolveColumnWidth(definition, options);
                columns.Add(new ExcelVisualColumn(column, x, width));
                x += width;
            }

            var rows = new List<ExcelVisualRow>();
            double y = 0D;
            for (int row = firstRow; row <= lastRow; row++) {
                rowDefinitions.TryGetValue(row, out ExcelRowSnapshot? definition);
                if (IsHiddenRow(row, rowDefinitions, defaultRowsHidden) && !options.IncludeHidden) {
                    continue;
                }

                double height = ResolveRowHeight(definition, options);
                rows.Add(new ExcelVisualRow(row, y, height));
                y += height;
            }

            if (!options.IncludeHidden) {
                ReportHiddenRowColumnOmissions(
                    sheet,
                    range,
                    firstRow,
                    firstColumn,
                    lastRow,
                    lastColumn,
                    rowDefinitions,
                    defaultRowsHidden,
                    columnDefinitions,
                    diagnostics);
            }

            Dictionary<int, ExcelVisualColumn> columnsByIndex = columns.ToDictionary(column => column.Index);
            Dictionary<int, ExcelVisualRow> rowsByIndex = rows.ToDictionary(row => row.Index);
            var coveredByMerge = new HashSet<string>(StringComparer.Ordinal);
            var mergeOrigins = new Dictionary<string, (double Width, double Height)>(StringComparer.Ordinal);
            var fullMergeGeometry = new HashSet<string>(StringComparer.Ordinal);
            var boundedMergeGeometry = new HashSet<string>(StringComparer.Ordinal);
            long remainingVisibleMergeCellWork = options.MaximumRenderedCells;
            long remainingFullMergeCellWork = options.MaximumRenderedCells;
            foreach (ExcelMergedRangeSnapshot merge in merges) {
                int visibleStartRow = Math.Max(firstRow, merge.StartRow);
                int visibleStartColumn = Math.Max(firstColumn, merge.StartColumn);
                int visibleEndRow = Math.Min(lastRow, merge.EndRow);
                int visibleEndColumn = Math.Min(lastColumn, merge.EndColumn);
                if (visibleStartRow > visibleEndRow || visibleStartColumn > visibleEndColumn) {
                    continue;
                }

                long visibleCellCount = ((long)visibleEndRow - visibleStartRow + 1L) *
                    ((long)visibleEndColumn - visibleStartColumn + 1L);
                long fullCellCount = ((long)merge.EndRow - merge.StartRow + 1L) *
                    ((long)merge.EndColumn - merge.StartColumn + 1L);
                if (visibleCellCount <= 0L || visibleCellCount > remainingVisibleMergeCellWork) {
                    continue;
                }

                remainingVisibleMergeCellWork -= visibleCellCount;
                boundedMergeGeometry.Add(MergeGeometryKey(merge));
                bool originIsVisible = merge.StartRow >= firstRow && merge.StartRow <= lastRow
                    && merge.StartColumn >= firstColumn && merge.StartColumn <= lastColumn
                    && rowsByIndex.ContainsKey(merge.StartRow)
                    && columnsByIndex.ContainsKey(merge.StartColumn);
                bool preserveFullGeometry = !originIsVisible
                    && fullCellCount > 0L
                    && fullCellCount <= remainingFullMergeCellWork;
                if (preserveFullGeometry) {
                    remainingFullMergeCellWork -= fullCellCount;
                    fullMergeGeometry.Add(MergeGeometryKey(merge));
                }

                double mergeWidth = columns
                    .Where(column => column.Index >= visibleStartColumn && column.Index <= visibleEndColumn)
                    .Sum(column => column.Width);
                double mergeHeight = rows
                    .Where(row => row.Index >= visibleStartRow && row.Index <= visibleEndRow)
                    .Sum(row => row.Height);
                mergeOrigins[Key(merge.StartRow, merge.StartColumn)] = (mergeWidth, mergeHeight);
                for (int row = visibleStartRow; row <= visibleEndRow; row++) {
                    for (int column = visibleStartColumn; column <= visibleEndColumn; column++) {
                        if (row != merge.StartRow || column != merge.StartColumn) {
                            coveredByMerge.Add(Key(row, column));
                        }
                    }
                }
            }

            var cells = new List<ExcelVisualCell>();
            Dictionary<string, ExcelHyperlinkSnapshot> hyperlinkMap = rows.Count > 0 && columns.Count > 0
                ? ExcelWorksheetHyperlinkResolver.BuildMap(
                    sheet.WorksheetPart,
                    rows.Min(row => row.Index),
                    columns.Min(column => column.Index),
                    rows.Max(row => row.Index),
                    columns.Max(column => column.Index))
                : new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelVisualRow row in rows) {
                foreach (ExcelVisualColumn column in columns) {
                    string key = Key(row.Index, column.Index);
                    bool covered = coveredByMerge.Contains(key);
                    ExcelHyperlinkSnapshot? hyperlink = null;
                    if (!covered) {
                        hyperlinkMap.TryGetValue(A1.CellReference(row.Index, column.Index), out hyperlink);
                    }

                    double width = column.Width;
                    double height = row.Height;
                    if (mergeOrigins.TryGetValue(key, out var mergeSize)) {
                        width = mergeSize.Width;
                        height = mergeSize.Height;
                    }

                    ExcelCellStyleSnapshot style = sheet.GetCellStyle(row.Index, column.Index);
                    ExcelCellData valueData = covered
                        ? new ExcelCellData(ExcelCellDataKind.Blank, null)
                        : sheet.GetCellValueSnapshot(row.Index, column.Index);
                    string rawText = sheet.TryGetCellText(row.Index, column.Index, out string cellText)
                        ? cellText
                        : string.Empty;
                    string text = string.IsNullOrEmpty(rawText)
                        ? string.Empty
                        : FormatCellDisplayText(rawText, style, sheet.Document.DateSystem);
                    ExcelVisualCellValueKind valueKind = ResolveVisualCellValueKind(valueData, rawText, style);
                    IReadOnlyList<ExcelVisualTextRun> richTextRuns = covered
                        ? Array.Empty<ExcelVisualTextRun>()
                        : BuildRichTextRuns(sheet.GetRichText(row.Index, column.Index));
                    cells.Add(new ExcelVisualCell(
                        row.Index,
                        column.Index,
                        column.X,
                        row.Y,
                        width,
                        height,
                        covered ? string.Empty : text,
                        style,
                        covered,
                        hyperlink,
                        richTextRuns,
                        valueKind));
                }
            }

            AddIntersectingMergeOriginCells(sheet, options, firstRow, firstColumn, lastRow, lastColumn, merges, boundedMergeGeometry, fullMergeGeometry, rowDefinitions, defaultRowsHidden, columnDefinitions, columnsByIndex, rowsByIndex, hyperlinkMap, cells);

            ExcelConditionalVisualState conditionalVisuals = options.IncludeConditionalFormatting
                ? ExcelConditionalVisualEvaluator.Evaluate(sheet, cells, range, options.ConditionalFormattingDate ?? DateTime.Today, diagnostics)
                : ExcelConditionalVisualState.Empty;
            if (conditionalVisuals.CellFormats.Count > 0) {
                cells = ApplyConditionalCellFormats(cells, conditionalVisuals.CellFormats);
            }

            List<ExcelVisualSparkline> sparklines = BuildSparklines(sheet, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, diagnostics);
            List<ExcelVisualDrawingObject> drawingObjects = BuildDrawingObjects(sheet, options, firstRow, firstColumn, lastRow, lastColumn, rowDefinitions, defaultRowsHidden, columnDefinitions, columnsByIndex, rowsByIndex, diagnostics);
            List<ExcelVisualImage> images = BuildImages(sheet, options, sourceImageBudget, firstRow, firstColumn, lastRow, lastColumn, rowDefinitions, defaultRowsHidden, columnDefinitions, columnsByIndex, rowsByIndex, diagnostics);
            List<ExcelVisualChart> charts = BuildCharts(sheet, options, firstRow, firstColumn, lastRow, lastColumn, rowDefinitions, defaultRowsHidden, columnDefinitions, columnsByIndex, rowsByIndex, diagnostics);
            IReadOnlyList<ExcelVisualBounds> commentBodyObstacles = BuildCommentBodyObstacles(drawingObjects, images, charts);
            CommentVisuals commentVisuals = BuildCommentVisuals(sheet, options, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, x, y, commentBodyObstacles, diagnostics);
            List<ExcelVisualDrawingLayer> drawingLayers = BuildDrawingLayers(drawingObjects, images, charts, commentVisuals.Bodies);

            return new ExcelRangeVisualSnapshot(
                sheet.Name,
                range,
                firstRow,
                firstColumn,
                lastRow,
                lastColumn,
                columns.AsReadOnly(),
                rows.AsReadOnly(),
                cells.AsReadOnly(),
                conditionalVisuals.DataBars,
                conditionalVisuals.Icons,
                commentVisuals.Indicators,
                commentVisuals.Bodies,
                sparklines.AsReadOnly(),
                drawingLayers.AsReadOnly(),
                drawingObjects.AsReadOnly(),
                images.AsReadOnly(),
                charts.AsReadOnly(),
                diagnostics.AsReadOnly());
        }

        private static List<ExcelVisualCell> ApplyConditionalCellFormats(IReadOnlyList<ExcelVisualCell> cells, IReadOnlyDictionary<string, ExcelConditionalCellFormat> formats) {
            var resolved = new List<ExcelVisualCell>(cells.Count);
            foreach (ExcelVisualCell cell in cells) {
                if (formats.TryGetValue(Key(cell.Row, cell.Column), out ExcelConditionalCellFormat? format)) {
                    resolved.Add(new ExcelVisualCell(
                        cell.Row,
                        cell.Column,
                        cell.X,
                        cell.Y,
                        cell.Width,
                        cell.Height,
                        cell.Text,
                        CloneStyleWithConditionalFormat(cell.Style, format),
                        cell.CoveredByMerge,
                        cell.Hyperlink,
                        cell.RichTextRuns,
                        cell.ValueKind));
                    continue;
                }

                resolved.Add(cell);
            }

            return resolved;
        }

        private static IReadOnlyList<ExcelVisualTextRun> BuildRichTextRuns(IReadOnlyList<ExcelRichTextRun> runs) {
            if (runs.Count == 0) {
                return Array.Empty<ExcelVisualTextRun>();
            }

            var visualRuns = new List<ExcelVisualTextRun>(runs.Count);
            foreach (ExcelRichTextRun run in runs) {
                if (string.IsNullOrEmpty(run.Text)) {
                    continue;
                }

                visualRuns.Add(new ExcelVisualTextRun(
                    run.Text,
                    run.Bold,
                    run.Italic,
                    run.Underline,
                    run.Strikethrough,
                    NormalizeRunColor(run.FontColor),
                    string.IsNullOrWhiteSpace(run.FontName) ? null : run.FontName,
                    run.FontSize));
            }

            return visualRuns.Count == 0 ? Array.Empty<ExcelVisualTextRun>() : visualRuns.AsReadOnly();
        }

        private static string? NormalizeRunColor(string? color) {
            if (string.IsNullOrWhiteSpace(color)) {
                return null;
            }

            string value = color!.Trim().TrimStart('#');
            if (value.Length == 6) {
                return "FF" + value.ToUpperInvariant();
            }

            return value.Length == 8 ? value.ToUpperInvariant() : null;
        }

        private static void AddIntersectingMergeOriginCells(
            ExcelSheet sheet,
            ExcelImageExportOptions options,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyList<ExcelMergedRangeSnapshot> merges,
            ISet<string> boundedMergeGeometry,
            ISet<string> fullMergeGeometry,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinkMap,
            List<ExcelVisualCell> cells) {
            foreach (ExcelMergedRangeSnapshot merge in merges) {
                if (!boundedMergeGeometry.Contains(MergeGeometryKey(merge))) {
                    continue;
                }

                if (merge.StartRow >= firstRow &&
                    merge.StartRow <= lastRow &&
                    merge.StartColumn >= firstColumn &&
                    merge.StartColumn <= lastColumn &&
                    rowsByIndex.ContainsKey(merge.StartRow) &&
                    columnsByIndex.ContainsKey(merge.StartColumn)) {
                    continue;
                }

                bool preserveFullGeometry = fullMergeGeometry.Contains(MergeGeometryKey(merge));
                double width;
                double height;
                double x;
                double y;
                if (preserveFullGeometry) {
                    width = 0D;
                    for (int column = merge.StartColumn; column <= merge.EndColumn; column++) {
                        width += ResolveVisibleColumnWidth(column, columnDefinitions, options);
                    }

                    height = 0D;
                    for (int row = merge.StartRow; row <= merge.EndRow; row++) {
                        height += ResolveVisibleRowHeight(row, rowDefinitions, defaultRowsHidden, options);
                    }

                    x = ResolveRelativeColumnOffset(firstColumn, merge.StartColumn, 0, columnDefinitions, options);
                    y = ResolveRelativeRowOffset(firstRow, merge.StartRow, 0, rowDefinitions, defaultRowsHidden, options);
                } else {
                    ExcelVisualColumn[] visibleColumns = columnsByIndex.Values
                        .Where(column => column.Index >= Math.Max(firstColumn, merge.StartColumn) && column.Index <= Math.Min(lastColumn, merge.EndColumn))
                        .OrderBy(column => column.Index)
                        .ToArray();
                    ExcelVisualRow[] visibleRows = rowsByIndex.Values
                        .Where(row => row.Index >= Math.Max(firstRow, merge.StartRow) && row.Index <= Math.Min(lastRow, merge.EndRow))
                        .OrderBy(row => row.Index)
                        .ToArray();
                    if (visibleColumns.Length == 0 || visibleRows.Length == 0) {
                        continue;
                    }

                    width = visibleColumns.Sum(column => column.Width);
                    height = visibleRows.Sum(row => row.Height);
                    x = visibleColumns[0].X;
                    y = visibleRows[0].Y;
                }

                if (width <= 0D || height <= 0D) {
                    continue;
                }

                ExcelCellStyleSnapshot style = sheet.GetCellStyle(merge.StartRow, merge.StartColumn);
                ExcelCellData valueData = sheet.GetCellValueSnapshot(merge.StartRow, merge.StartColumn);
                string rawText = sheet.TryGetCellText(merge.StartRow, merge.StartColumn, out string cellText)
                    ? cellText
                    : string.Empty;
                string text = string.IsNullOrEmpty(rawText)
                    ? string.Empty
                    : FormatCellDisplayText(rawText, style, sheet.Document.DateSystem);
                hyperlinkMap.TryGetValue(A1.CellReference(merge.StartRow, merge.StartColumn), out ExcelHyperlinkSnapshot? hyperlink);
                cells.Add(new ExcelVisualCell(
                    merge.StartRow,
                    merge.StartColumn,
                    x,
                    y,
                    width,
                    height,
                    text,
                    style,
                    coveredByMerge: false,
                    hyperlink,
                    BuildRichTextRuns(sheet.GetRichText(merge.StartRow, merge.StartColumn)),
                    ResolveVisualCellValueKind(valueData, rawText, style)));
            }
        }

        private static List<ExcelVisualDrawingObject> BuildDrawingObjects(
            ExcelSheet sheet,
            ExcelImageExportOptions options,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var drawingObjects = new List<ExcelVisualDrawingObject>();
            if (!options.IncludeDrawingObjects) {
                return drawingObjects;
            }

            foreach (ExcelWorksheetDrawingObjectInfo drawing in ExcelWorksheetDrawingObjectResolver.FindDrawingObjects(sheet.WorksheetPart)) {
                if (!drawing.IsRenderable) {
                    ReportUnsupportedDrawingObject(sheet, drawing, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, diagnostics);
                    continue;
                }

                if (!options.IncludeHidden && IsHiddenAnchorInRange(drawing.Row, drawing.Column, firstRow, firstColumn, lastRow, lastColumn, rowDefinitions, defaultRowsHidden, columnDefinitions)) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.DrawingShapeAnchorHidden,
                        "Worksheet drawing object is anchored to a hidden row or column and was omitted from the image export.",
                        GetDrawingDiagnosticSource(sheet, drawing)));
                    continue;
                }

                if (!TryGetImageAnchor(
                    options,
                    firstRow,
                    firstColumn,
                    rowDefinitions,
                    defaultRowsHidden,
                    columnDefinitions,
                    columnsByIndex,
                    rowsByIndex,
                    drawing.Column,
                    drawing.Row,
                    drawing.OffsetXPixels,
                    drawing.OffsetYPixels,
                    drawing.WidthPixels,
                    drawing.HeightPixels,
                    drawing.ToColumn,
                    drawing.ToRow,
                    drawing.ToOffsetXPixels,
                    drawing.ToOffsetYPixels,
                    out double x,
                    out double y,
                    out double width,
                    out double height)) {
                    continue;
                }

                if (!drawing.ShapeKind.HasValue) {
                    continue;
                }

                drawingObjects.Add(new ExcelVisualDrawingObject(
                    drawing.Name,
                    drawing.Order,
                    drawing.ShapePresetName,
                    drawing.ShapeKind.Value,
                    drawing.HorizontalFlip,
                    drawing.VerticalFlip,
                    drawing.RotationDegrees,
                    x,
                    y,
                    width,
                    height,
                    drawing.FillColorArgb,
                    drawing.StrokeColorArgb,
                    drawing.StrokeWidth,
                    drawing.StrokeDashStyle,
                    drawing.StrokeLineCap,
                    drawing.StrokeLineJoin,
                    drawing.Text,
                    drawing.TextAlignment,
                    drawing.TextVerticalAlignment,
                    drawing.TextColorArgb,
                    drawing.TextFontFamily,
                    drawing.TextFontSize,
                    drawing.TextFontStyle,
                    drawing.TextWrap,
                    drawing.TextShrinkToFit,
                    drawing.TextResizeShapeToFit,
                    drawing.TextOrientation,
                    drawing.TextInsetLeft,
                    drawing.TextInsetTop,
                    drawing.TextInsetRight,
                    drawing.TextInsetBottom,
                    drawing.Glow,
                    drawing.Shadow,
                    GetDrawingDiagnosticSource(sheet, drawing)));
            }

            return drawingObjects;
        }

        private static ExcelCellStyleSnapshot CloneStyleWithConditionalFormat(ExcelCellStyleSnapshot style, ExcelConditionalCellFormat format) {
            string? fillColorArgb = string.IsNullOrWhiteSpace(format.FillColorArgb) ? null : format.FillColorArgb;
            return new ExcelCellStyleSnapshot {
                StyleIndex = style.StyleIndex,
                NumberFormatId = style.NumberFormatId,
                NumberFormatCode = style.NumberFormatCode,
                IsDateLike = style.IsDateLike,
                Bold = format.FontBold ?? style.Bold,
                Italic = format.FontItalic ?? style.Italic,
                Underline = format.FontUnderline ?? style.Underline,
                FontName = string.IsNullOrWhiteSpace(format.FontName) ? style.FontName : format.FontName,
                IsFontFamilyExplicit = !string.IsNullOrWhiteSpace(format.FontName) || style.IsFontFamilyExplicit,
                FontSize = format.FontSize ?? style.FontSize,
                TextRotation = style.TextRotation,
                FontColorArgb = string.IsNullOrWhiteSpace(format.FontColorArgb) ? style.FontColorArgb : format.FontColorArgb,
                FillColorArgb = fillColorArgb ?? style.FillColorArgb,
                FillPatternType = fillColorArgb == null ? style.FillPatternType : "solid",
                FillPatternForegroundColorArgb = fillColorArgb ?? style.FillPatternForegroundColorArgb,
                FillPatternBackgroundColorArgb = fillColorArgb ?? style.FillPatternBackgroundColorArgb,
                FillGradientUnsupported = fillColorArgb == null && style.FillGradientUnsupported,
                FillGradientStartColorArgb = fillColorArgb == null ? style.FillGradientStartColorArgb : null,
                FillGradientEndColorArgb = fillColorArgb == null ? style.FillGradientEndColorArgb : null,
                FillGradientStops = fillColorArgb == null ? style.FillGradientStops : Array.Empty<ExcelGradientFillStopSnapshot>(),
                FillGradientDegree = fillColorArgb == null ? style.FillGradientDegree : null,
                Border = MergeConditionalBorder(style.Border, format.Border),
                HorizontalAlignment = style.HorizontalAlignment,
                VerticalAlignment = style.VerticalAlignment,
                TextIndent = style.TextIndent,
                WrapText = style.WrapText,
                ShrinkToFit = style.ShrinkToFit
            };
        }

        private static ExcelCellBorderSnapshot? MergeConditionalBorder(ExcelCellBorderSnapshot? styleBorder, ExcelCellBorderSnapshot? conditionalBorder) {
            if (conditionalBorder == null) {
                return styleBorder;
            }

            if (styleBorder == null) {
                return conditionalBorder;
            }

            return new ExcelCellBorderSnapshot {
                Left = conditionalBorder.Left ?? styleBorder.Left,
                Right = conditionalBorder.Right ?? styleBorder.Right,
                Top = conditionalBorder.Top ?? styleBorder.Top,
                Bottom = conditionalBorder.Bottom ?? styleBorder.Bottom,
                Diagonal = conditionalBorder.Diagonal ?? styleBorder.Diagonal,
                DiagonalUp = conditionalBorder.Diagonal != null ? conditionalBorder.DiagonalUp : styleBorder.DiagonalUp,
                DiagonalDown = conditionalBorder.Diagonal != null ? conditionalBorder.DiagonalDown : styleBorder.DiagonalDown
            };
        }

        private static List<ExcelVisualImage> BuildImages(
            ExcelSheet sheet,
            ExcelImageExportOptions options,
            ExcelSourceImageBudget sourceImageBudget,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var images = new List<ExcelVisualImage>();
            if (!options.IncludeImages) {
                return images;
            }

            foreach (ExcelImage image in sheet.Images) {
                bool absoluteAnchor = image.TryGetAbsoluteAnchorBounds(out int absoluteX, out int absoluteY, out int absoluteWidth, out int absoluteHeight);
                if (!absoluteAnchor && !options.IncludeHidden && IsHiddenAnchorInRange(image.RowIndex, image.ColumnIndex, firstRow, firstColumn, lastRow, lastColumn, rowDefinitions, defaultRowsHidden, columnDefinitions)) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.ImageAnchorHidden,
                        "Worksheet image is anchored to a hidden row or column and was omitted from the image export.",
                        GetImageDiagnosticSource(sheet, image)));
                    continue;
                }

                if (!TryGetImageAnchor(
                    options,
                    firstRow,
                    firstColumn,
                    rowDefinitions,
                    defaultRowsHidden,
                    columnDefinitions,
                    columnsByIndex,
                    rowsByIndex,
                    absoluteAnchor ? 1 : image.ColumnIndex,
                    absoluteAnchor ? 1 : image.RowIndex,
                    absoluteAnchor ? absoluteX : image.OffsetXPixels,
                    absoluteAnchor ? absoluteY : image.OffsetYPixels,
                    absoluteAnchor ? absoluteWidth : image.WidthPixels,
                    absoluteAnchor ? absoluteHeight : image.HeightPixels,
                    absoluteAnchor ? null : image.ToColumnIndex,
                    absoluteAnchor ? null : image.ToRowIndex,
                    absoluteAnchor ? 0 : image.ToOffsetXPixels,
                    absoluteAnchor ? 0 : image.ToOffsetYPixels,
                    out double x,
                    out double y,
                    out double width,
                    out double height)) {
                    continue;
                }

                string source = GetImageDiagnosticSource(sheet, image);
                if (!image.TryReadBytes(sourceImageBudget.RemainingBytes, out byte[] bytes)) {
                    const string message = "Worksheet image was omitted because its bytes could not be read within the remaining aggregate source-image budget.";
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(OfficeImageExportDiagnosticSeverity.Warning, ExcelImageExportDiagnosticCodes.ImageBytesMissing, message, source));
                    continue;
                }

                sourceImageBudget.Consume(bytes.LongLength);

                bool identifiedImage = OfficeImageReader.TryIdentify(bytes, image.Name, out OfficeImageInfo info);
                OfficeImageFormat detectedFormat = identifiedImage
                    ? info.Format
                    : OfficeImageFormat.Unknown;
                if (detectedFormat == OfficeImageFormat.Unknown) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.ImageFormatUnknown,
                        "Worksheet image bytes do not contain a recognized image header. Declared content type: '" + image.ContentType + "'. Export uses the caller codec when available and otherwise renders a visible fallback.",
                        source));
                }

                images.Add(new ExcelVisualImage(
                    image.Name,
                    image.DrawingOrder,
                    image.ContentType,
                    detectedFormat,
                    bytes,
                    identifiedImage ? info.Width : 0D,
                    identifiedImage ? info.Height : 0D,
                    x,
                    y,
                    width,
                    height,
                    image.CropLeftRatio,
                    image.CropTopRatio,
                    image.CropRightRatio,
                    image.CropBottomRatio,
                    image.RotationDegrees,
                    image.FlipHorizontal,
                    image.FlipVertical,
                    source));
            }

            return images;
        }

        private static List<ExcelVisualChart> BuildCharts(
            ExcelSheet sheet,
            ExcelImageExportOptions options,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var charts = new List<ExcelVisualChart>();
            if (!options.IncludeCharts) {
                return charts;
            }

            foreach (ExcelChart chart in sheet.Charts) {
                if (!chart.TryGetSnapshot(out ExcelChartSnapshot chartSnapshot)) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(OfficeImageExportDiagnosticSeverity.Warning, ExcelImageExportDiagnosticCodes.ChartSnapshotUnavailable, "Worksheet chart data could not be converted to a renderable snapshot.", sheet.Name + "!" + chart.Name));
                    continue;
                }

                foreach (OfficeImageExportDiagnostic diagnostic in chartSnapshot.Diagnostics) {
                    diagnostics.Add(diagnostic);
                }

                bool absoluteAnchor = chart.TryGetAbsoluteAnchorBounds(out int absoluteX, out int absoluteY, out int absoluteWidth, out int absoluteHeight);
                if (!absoluteAnchor && !options.IncludeHidden && IsHiddenAnchorInRange(chartSnapshot.RowIndex, chartSnapshot.ColumnIndex, firstRow, firstColumn, lastRow, lastColumn, rowDefinitions, defaultRowsHidden, columnDefinitions)) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.ChartAnchorHidden,
                        "Worksheet chart is anchored to a hidden row or column and was omitted from the image export.",
                        sheet.Name + "!" + chartSnapshot.Name));
                    continue;
                }

                if (!TryGetImageAnchor(
                    options,
                    firstRow,
                    firstColumn,
                    rowDefinitions,
                    defaultRowsHidden,
                    columnDefinitions,
                    columnsByIndex,
                    rowsByIndex,
                    absoluteAnchor ? 1 : chartSnapshot.ColumnIndex,
                    absoluteAnchor ? 1 : chartSnapshot.RowIndex,
                    absoluteAnchor ? absoluteX : chartSnapshot.OffsetXPixels,
                    absoluteAnchor ? absoluteY : chartSnapshot.OffsetYPixels,
                    absoluteAnchor ? absoluteWidth : chartSnapshot.WidthPixels,
                    absoluteAnchor ? absoluteHeight : chartSnapshot.HeightPixels,
                    toColumnIndex: null,
                    toRowIndex: null,
                    toOffsetXPixels: 0,
                    toOffsetYPixels: 0,
                    out double x,
                    out double y,
                    out double width,
                    out double height)) {
                    continue;
                }

                charts.Add(new ExcelVisualChart(chartSnapshot, chart.DrawingOrder, x, y, width, height));
            }

            return charts;
        }

        private static List<ExcelVisualDrawingLayer> BuildDrawingLayers(
            IReadOnlyList<ExcelVisualDrawingObject> drawingObjects,
            IReadOnlyList<ExcelVisualImage> images,
            IReadOnlyList<ExcelVisualChart> charts,
            IReadOnlyList<ExcelVisualCommentBody> commentBodies) {
            var layers = new List<ExcelVisualDrawingLayer>(drawingObjects.Count + images.Count + charts.Count + commentBodies.Count);
            foreach (ExcelVisualDrawingObject drawingObject in drawingObjects) {
                layers.Add(ExcelVisualDrawingLayer.FromDrawingObject(drawingObject));
            }

            foreach (ExcelVisualImage image in images) {
                layers.Add(ExcelVisualDrawingLayer.FromImage(image));
            }

            foreach (ExcelVisualChart chart in charts) {
                layers.Add(ExcelVisualDrawingLayer.FromChart(chart));
            }

            int commentOrder = layers.Count == 0 ? 0 : layers.Max(layer => layer.Order) + 1;
            foreach (ExcelVisualCommentBody commentBody in commentBodies) {
                layers.Add(ExcelVisualDrawingLayer.FromCommentBody(commentBody, commentOrder));
                commentOrder++;
            }

            return layers
                .OrderBy(layer => layer.Order)
                .ThenBy(layer => (int)layer.Kind)
                .ToList();
        }

        private static IReadOnlyList<ExcelVisualBounds> BuildCommentBodyObstacles(
            IReadOnlyList<ExcelVisualDrawingObject> drawingObjects,
            IReadOnlyList<ExcelVisualImage> images,
            IReadOnlyList<ExcelVisualChart> charts) {
            var bounds = new List<ExcelVisualBounds>(drawingObjects.Count + images.Count + charts.Count);
            foreach (ExcelVisualDrawingObject drawingObject in drawingObjects) {
                AddCommentBodyObstacle(bounds, drawingObject.X, drawingObject.Y, drawingObject.Width, drawingObject.Height);
            }

            foreach (ExcelVisualImage image in images) {
                AddCommentBodyObstacle(bounds, image.X, image.Y, image.Width, image.Height);
            }

            foreach (ExcelVisualChart chart in charts) {
                AddCommentBodyObstacle(bounds, chart.X, chart.Y, chart.Width, chart.Height);
            }

            return bounds.AsReadOnly();
        }

        private static void AddCommentBodyObstacle(List<ExcelVisualBounds> bounds, double x, double y, double width, double height) {
            var value = new ExcelVisualBounds(x, y, width, height);
            if (!value.IsEmpty) {
                bounds.Add(value);
            }
        }

        private static CommentVisuals BuildCommentVisuals(
            ExcelSheet sheet,
            ExcelImageExportOptions options,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            double snapshotWidth,
            double snapshotHeight,
            IReadOnlyList<ExcelVisualBounds> bodyObstacles,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var indicators = new Dictionary<string, ExcelVisualCommentIndicator>(StringComparer.OrdinalIgnoreCase);
            var bodies = new List<ExcelVisualCommentBody>();
            Dictionary<string, ExcelCommentSnapshot> comments = ExcelWorksheetCommentResolver.BuildLegacyCommentMap(sheet.WorksheetPart);
            foreach (KeyValuePair<string, ExcelCommentSnapshot> pair in comments.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                string reference = pair.Key;
                if (!TryResolveVisibleExportCell(reference, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, out string sourceCell)) {
                    continue;
                }

                string source = sheet.Name + "!" + sourceCell;
                if (TryCreateCommentIndicator(sheet.Name, sourceCell, threaded: false, columnsByIndex, rowsByIndex, out ExcelVisualCommentIndicator indicator)) {
                    indicators[sourceCell] = indicator;
                    if (options.ShowCommentBodies && TryCreateCommentBody(
                        indicator,
                        snapshotWidth,
                        snapshotHeight,
                        pair.Value.Author,
                        pair.Value.Text,
                        BuildRichTextRuns(pair.Value.RichTextRuns),
                        bodyObstacles,
                        out ExcelVisualCommentBody body)) {
                        bodies.Add(body);
                    }
                }

                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    options.ShowCommentBodies
                        ? ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation
                        : ExcelImageExportDiagnosticCodes.CellCommentUnsupported,
                    options.ShowCommentBodies
                        ? "Excel cell comment/note body is rendered as a dependency-free callout approximation, not an Excel-exact popover."
                        : "Excel cell comment/note body is not rendered by the dependency-free image exporter yet; only the cell indicator is rendered.",
                    source));
            }

            WorkbookPart? workbookPart = sheet.WorksheetPart.OpenXmlPackage is SpreadsheetDocument spreadsheetDocument
                ? spreadsheetDocument.WorkbookPart
                : null;
            IReadOnlyDictionary<string, string> people = workbookPart == null
                ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                : ExcelWorksheetCommentResolver.BuildThreadedCommentPersonMap(workbookPart);
            Dictionary<string, List<ExcelThreadedCommentSnapshot>> threadedComments = ExcelWorksheetCommentResolver.BuildThreadedCommentMap(sheet.WorksheetPart, people);
            foreach (KeyValuePair<string, List<ExcelThreadedCommentSnapshot>> pair in threadedComments.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                if (!TryResolveVisibleExportCell(pair.Key, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, out string sourceCell)) {
                    continue;
                }

                string source = sheet.Name + "!" + sourceCell;
                string unsupportedMessage = pair.Value.Count == 1
                    ? "Excel threaded comment body is not rendered by the dependency-free image exporter yet; only the cell indicator is rendered."
                    : pair.Value.Count.ToString(CultureInfo.InvariantCulture) + " Excel threaded comment bodies are not rendered by the dependency-free image exporter yet; only the cell indicator is rendered.";
                if (TryCreateCommentIndicator(sheet.Name, sourceCell, threaded: true, columnsByIndex, rowsByIndex, out ExcelVisualCommentIndicator indicator)) {
                    indicators[sourceCell] = indicator;
                    if (options.ShowCommentBodies && TryCreateCommentBody(
                        indicator,
                        snapshotWidth,
                        snapshotHeight,
                        ResolveThreadedCommentTitle(pair.Value),
                        FormatThreadedCommentBody(pair.Value),
                        Array.Empty<ExcelVisualTextRun>(),
                        bodyObstacles,
                        out ExcelVisualCommentBody body)) {
                        bodies.Add(body);
                    }
                }

                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    options.ShowCommentBodies
                        ? ExcelImageExportDiagnosticCodes.ThreadedCommentBodyApproximation
                        : ExcelImageExportDiagnosticCodes.ThreadedCommentUnsupported,
                    options.ShowCommentBodies
                        ? "Excel threaded comment body is rendered as a dependency-free callout approximation, not an Excel-exact popover."
                        : unsupportedMessage,
                    source));
            }

            return new CommentVisuals(
                indicators.Values
                    .OrderBy(indicator => indicator.Row)
                    .ThenBy(indicator => indicator.Column)
                    .ToList()
                    .AsReadOnly(),
                bodies
                    .OrderBy(body => body.Row)
                    .ThenBy(body => body.Column)
                    .ToList()
                    .AsReadOnly());
        }

        private static bool TryCreateCommentIndicator(
            string sheetName,
            string sourceCell,
            bool threaded,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            out ExcelVisualCommentIndicator indicator) {
            indicator = null!;
            if (!A1.TryParseCellReferenceFast(sourceCell, out int row, out int column) ||
                !columnsByIndex.TryGetValue(column, out ExcelVisualColumn? visualColumn) ||
                !rowsByIndex.TryGetValue(row, out ExcelVisualRow? visualRow)) {
                return false;
            }

            indicator = new ExcelVisualCommentIndicator(
                row,
                column,
                visualColumn.X,
                visualRow.Y,
                visualColumn.Width,
                visualRow.Height,
                threaded,
                sheetName + "!" + sourceCell);
            return true;
        }

        private static bool TryCreateCommentBody(
            ExcelVisualCommentIndicator indicator,
            double snapshotWidth,
            double snapshotHeight,
            string? title,
            string? text,
            IReadOnlyList<ExcelVisualTextRun>? richTextRuns,
            IReadOnlyList<ExcelVisualBounds> bodyObstacles,
            out ExcelVisualCommentBody body) {
            body = null!;
            if (snapshotWidth <= 0D || snapshotHeight <= 0D) {
                return false;
            }

            string resolvedTitle = string.IsNullOrWhiteSpace(title)
                ? indicator.Threaded ? "Threaded comment" : "Comment"
                : title!.Trim();
            string resolvedText = string.IsNullOrWhiteSpace(text) ? "(empty)" : NormalizeCommentBodyText(text!);
            double width = Math.Min(220D, Math.Max(132D, snapshotWidth * 0.46D));
            width = Math.Min(width, Math.Max(96D, snapshotWidth - 8D));
            double height = EstimateCommentBodyHeight(resolvedTitle, resolvedText);
            height = Math.Min(height, Math.Max(48D, snapshotHeight - 8D));

            ExcelVisualBounds placement = ChooseCommentBodyPlacement(indicator, snapshotWidth, snapshotHeight, width, height, bodyObstacles);

            body = new ExcelVisualCommentBody(
                indicator.Row,
                indicator.Column,
                placement.X,
                placement.Y,
                width,
                height,
                indicator.X + indicator.Width,
                indicator.Y + Math.Min(indicator.Height, 8D),
                indicator.Threaded,
                resolvedTitle,
                resolvedText,
                richTextRuns,
                indicator.Source);
            return true;
        }

        private static ExcelVisualBounds ChooseCommentBodyPlacement(
            ExcelVisualCommentIndicator indicator,
            double snapshotWidth,
            double snapshotHeight,
            double width,
            double height,
            IReadOnlyList<ExcelVisualBounds> bodyObstacles) {
            double rightX = indicator.X + indicator.Width + 6D;
            double leftX = indicator.X - width - 6D;
            double sideY = indicator.Y + 2D;

            var candidates = new[] {
                CreateCommentBodyCandidate(rightX, sideY, width, height, snapshotWidth, snapshotHeight, 0),
                CreateCommentBodyCandidate(leftX, sideY, width, height, snapshotWidth, snapshotHeight, 1),
                CreateCommentBodyCandidate(rightX, indicator.Y + indicator.Height + 6D, width, height, snapshotWidth, snapshotHeight, 2),
                CreateCommentBodyCandidate(leftX, indicator.Y + indicator.Height + 6D, width, height, snapshotWidth, snapshotHeight, 3),
                CreateCommentBodyCandidate(indicator.X, indicator.Y + indicator.Height + 6D, width, height, snapshotWidth, snapshotHeight, 4),
                CreateCommentBodyCandidate(indicator.X, indicator.Y - height - 6D, width, height, snapshotWidth, snapshotHeight, 5)
            };

            CommentBodyPlacementCandidate best = candidates[0];
            double bestScore = ScoreCommentBodyPlacement(best.Bounds, bodyObstacles, best.Preference);
            for (int i = 1; i < candidates.Length; i++) {
                double score = ScoreCommentBodyPlacement(candidates[i].Bounds, bodyObstacles, candidates[i].Preference);
                if (score < bestScore) {
                    best = candidates[i];
                    bestScore = score;
                }
            }

            return best.Bounds;
        }

        private static CommentBodyPlacementCandidate CreateCommentBodyCandidate(
            double x,
            double y,
            double width,
            double height,
            double snapshotWidth,
            double snapshotHeight,
            int preference) {
            double resolvedX = ClampCommentBodyCoordinate(x, width, snapshotWidth);
            double resolvedY = ClampCommentBodyCoordinate(y, height, snapshotHeight);
            return new CommentBodyPlacementCandidate(new ExcelVisualBounds(resolvedX, resolvedY, width, height), preference);
        }

        private static double ClampCommentBodyCoordinate(double value, double length, double availableLength) {
            double minimum = 4D;
            double maximum = Math.Max(minimum, availableLength - length - 4D);
            if (value < minimum) {
                return minimum;
            }

            return value > maximum ? maximum : value;
        }

        private static double ScoreCommentBodyPlacement(ExcelVisualBounds bounds, IReadOnlyList<ExcelVisualBounds> bodyObstacles, int preference) {
            double overlapArea = 0D;
            foreach (ExcelVisualBounds obstacle in bodyObstacles) {
                overlapArea += bounds.IntersectionArea(obstacle);
            }

            return overlapArea + (preference * 0.001D);
        }

        private static double EstimateCommentBodyHeight(string title, string text) {
            int lines = 1;
            foreach (string line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
                lines += Math.Max(1, (int)Math.Ceiling(line.Length / 34D));
            }

            if (!string.IsNullOrWhiteSpace(title)) {
                lines++;
            }

            return Math.Min(150D, Math.Max(58D, 18D + (lines * 13D)));
        }

        private static string ResolveThreadedCommentTitle(IReadOnlyList<ExcelThreadedCommentSnapshot> comments) {
            string? author = comments.Select(comment => comment.Author).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            return string.IsNullOrWhiteSpace(author)
                ? comments.Count == 1 ? "Threaded comment" : "Threaded comments"
                : author!;
        }

        private static string FormatThreadedCommentBody(IReadOnlyList<ExcelThreadedCommentSnapshot> comments) {
            var builder = new System.Text.StringBuilder();
            for (int i = 0; i < comments.Count; i++) {
                ExcelThreadedCommentSnapshot comment = comments[i];
                if (i > 0) {
                    builder.Append('\n');
                }

                string author = string.IsNullOrWhiteSpace(comment.Author) ? "Comment" : comment.Author!;
                string text = string.IsNullOrWhiteSpace(comment.Text) ? "(empty)" : NormalizeCommentBodyText(comment.Text);
                builder.Append(author).Append(": ").Append(text);
            }

            return builder.ToString();
        }

        private static string NormalizeCommentBodyText(string text) =>
            text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();

        private readonly struct CommentBodyPlacementCandidate {
            internal CommentBodyPlacementCandidate(ExcelVisualBounds bounds, int preference) {
                Bounds = bounds;
                Preference = preference;
            }

            internal ExcelVisualBounds Bounds { get; }

            internal int Preference { get; }
        }

        private sealed class CommentVisuals {
            internal CommentVisuals(IReadOnlyList<ExcelVisualCommentIndicator> indicators, IReadOnlyList<ExcelVisualCommentBody> bodies) {
                Indicators = indicators;
                Bodies = bodies;
            }

            internal IReadOnlyList<ExcelVisualCommentIndicator> Indicators { get; }

            internal IReadOnlyList<ExcelVisualCommentBody> Bodies { get; }
        }

        private static List<ExcelVisualSparkline> BuildSparklines(
            ExcelSheet sheet,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var visuals = new List<ExcelVisualSparkline>();
            IReadOnlyList<ResolvedSparkline> resolvedSparklines = ResolveSparklines(
                sheet,
                firstRow,
                firstColumn,
                lastRow,
                lastColumn,
                columnsByIndex,
                rowsByIndex);
            IReadOnlyDictionary<int, SparklineScaleRange> scaleRanges = ResolveSparklineScaleRanges(resolvedSparklines);
            foreach (ResolvedSparkline resolvedSparkline in resolvedSparklines) {
                ExcelWorksheetSparklineInfo sparkline = resolvedSparkline.Info;
                if (!TryResolveVisibleExportCell(sparkline.CellReference, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, out string sourceCell)
                    || !A1.TryParseCellReferenceFast(sourceCell, out int row, out int column)
                    || !rowsByIndex.TryGetValue(row, out ExcelVisualRow? visualRow)
                    || !columnsByIndex.TryGetValue(column, out ExcelVisualColumn? visualColumn)) {
                    continue;
                }

                string source = sheet.Name + "!" + sourceCell;
                if (!IsSupportedSparklineKind(sparkline.Kind)) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.SparklineKindUnsupported,
                        "Worksheet sparkline kind '" + sparkline.Kind + "' is not rendered by the dependency-free image exporter yet.",
                        source));
                    continue;
                }

                if (!resolvedSparkline.HasResolvedRange) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        resolvedSparkline.ExternalRange ? ExcelImageExportDiagnosticCodes.SparklineExternalRangeUnsupported : ExcelImageExportDiagnosticCodes.SparklineRangeUnsupported,
                        resolvedSparkline.ExternalRange
                            ? "Worksheet sparkline references data on another sheet; cross-sheet sparkline image rendering is not supported yet."
                            : "Worksheet sparkline data range could not be resolved by the dependency-free image exporter.",
                        source));
                    continue;
                }

                if (resolvedSparkline.Values.Count == 0) {
                    diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.SparklineDataMissing,
                        "Worksheet sparkline has no numeric values available for image rendering.",
                        source));
                    continue;
                }

                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation,
                    "Worksheet sparkline is rendered as a dependency-free approximation; Excel-specific date axis, hidden-data, and empty-cell behavior may differ.",
                    source));
                scaleRanges.TryGetValue(sparkline.GroupIndex, out SparklineScaleRange scaleRange);

                visuals.Add(new ExcelVisualSparkline(
                    row,
                    column,
                    visualColumn.X,
                    visualRow.Y,
                    visualColumn.Width,
                    visualRow.Height,
                    sparkline.Kind,
                    resolvedSparkline.Values,
                    sparkline.DisplayMarkers,
                    sparkline.DisplayHigh,
                    sparkline.DisplayLow,
                    sparkline.DisplayFirst,
                    sparkline.DisplayLast,
                    sparkline.DisplayNegative,
                    sparkline.DisplayAxis,
                    sparkline.SeriesColorArgb,
                    sparkline.AxisColorArgb,
                    sparkline.NegativeColorArgb,
                    sparkline.MarkersColorArgb,
                    sparkline.HighColorArgb,
                    sparkline.LowColorArgb,
                    sparkline.FirstColorArgb,
                    sparkline.LastColorArgb,
                    scaleRange.Minimum,
                    scaleRange.Maximum,
                    source));
            }

            return visuals;
        }

        private static IReadOnlyList<ResolvedSparkline> ResolveSparklines(
            ExcelSheet sheet,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex) {
            IReadOnlyList<ExcelWorksheetSparklineInfo> sparklines = ExcelWorksheetSparklineResolver.FindSparklines(sheet.WorksheetPart);
            var visibleIndexes = new HashSet<int>();
            var visibleGroups = new HashSet<int>();
            for (int index = 0; index < sparklines.Count; index++) {
                ExcelWorksheetSparklineInfo sparkline = sparklines[index];
                if (IsSupportedSparklineKind(sparkline.Kind) &&
                    TryResolveVisibleExportCell(
                        sparkline.CellReference,
                        firstRow,
                        firstColumn,
                        lastRow,
                        lastColumn,
                        columnsByIndex,
                        rowsByIndex,
                        out _)) {
                    visibleIndexes.Add(index);
                    visibleGroups.Add(sparkline.GroupIndex);
                }
            }

            var resolved = new ResolvedSparkline[sparklines.Count];
            long remainingSparklineDataCells = MaxSparklineDataCells;
            IEnumerable<int> resolutionOrder = Enumerable.Range(0, sparklines.Count)
                .Where(visibleIndexes.Contains)
                .Concat(Enumerable.Range(0, sparklines.Count).Where(index => !visibleIndexes.Contains(index)));
            foreach (int index in resolutionOrder) {
                ExcelWorksheetSparklineInfo sparkline = sparklines[index];
                if (!IsSupportedSparklineKind(sparkline.Kind) ||
                    !visibleGroups.Contains(sparkline.GroupIndex)) {
                    resolved[index] = new ResolvedSparkline(sparkline, Array.Empty<double>(), hasResolvedRange: false, externalRange: false);
                    continue;
                }

                if (!TryResolveSparklineDataRange(sheet.Name, sparkline.Formula, out string? dataRange, out bool externalRange) || dataRange == null) {
                    resolved[index] = new ResolvedSparkline(sparkline, Array.Empty<double>(), hasResolvedRange: false, externalRange);
                    continue;
                }

                if (!TryReadSparklineValues(sheet, dataRange, ref remainingSparklineDataCells, out IReadOnlyList<double> values)) {
                    resolved[index] = new ResolvedSparkline(sparkline, Array.Empty<double>(), hasResolvedRange: false, externalRange: false);
                    continue;
                }

                resolved[index] = new ResolvedSparkline(sparkline, values, hasResolvedRange: true, externalRange: false);
            }

            return Array.AsReadOnly(resolved);
        }

        private static IReadOnlyDictionary<int, SparklineScaleRange> ResolveSparklineScaleRanges(IReadOnlyList<ResolvedSparkline> sparklines) {
            var ranges = new Dictionary<int, SparklineScaleRange>();
            foreach (IGrouping<int, ResolvedSparkline> group in sparklines
                .Where(sparkline => sparkline.HasResolvedRange && sparkline.Values.Count > 0)
                .GroupBy(sparkline => sparkline.Info.GroupIndex)) {
                double minimum = group.SelectMany(sparkline => sparkline.Values).Min();
                double maximum = group.SelectMany(sparkline => sparkline.Values).Max();
                ranges[group.Key] = new SparklineScaleRange(minimum, maximum);
            }

            return ranges;
        }

        private readonly struct ResolvedSparkline {
            internal ResolvedSparkline(ExcelWorksheetSparklineInfo info, IReadOnlyList<double> values, bool hasResolvedRange, bool externalRange) {
                Info = info;
                Values = values;
                HasResolvedRange = hasResolvedRange;
                ExternalRange = externalRange;
            }

            internal ExcelWorksheetSparklineInfo Info { get; }

            internal IReadOnlyList<double> Values { get; }

            internal bool HasResolvedRange { get; }

            internal bool ExternalRange { get; }
        }

        private readonly struct SparklineScaleRange {
            internal SparklineScaleRange(double minimum, double maximum) {
                Minimum = minimum;
                Maximum = maximum;
            }

            internal double Minimum { get; }

            internal double Maximum { get; }
        }

        private static bool IsSupportedSparklineKind(string kind) {
            string normalized = NormalizeSparklineKind(kind);
            return normalized.Length == 0 || normalized == "line" || normalized == "column" || normalized == "stacked" || normalized == "winloss";
        }

        private static string NormalizeSparklineKind(string kind) =>
            (kind ?? string.Empty).Replace("-", string.Empty).Trim().ToLowerInvariant();

        private static bool TryResolveSparklineDataRange(string sheetName, string formula, out string? dataRange, out bool externalRange) {
            dataRange = null;
            externalRange = false;
            string value = (formula ?? string.Empty).Trim();
            if (value.StartsWith("=", StringComparison.Ordinal)) {
                value = value.Substring(1).Trim();
            }

            int separator = value.LastIndexOf('!');
            if (separator >= 0) {
                string formulaSheet = UnquoteSheetName(value.Substring(0, separator));
                value = value.Substring(separator + 1);
                if (!string.Equals(formulaSheet, sheetName, StringComparison.OrdinalIgnoreCase)) {
                    externalRange = true;
                    return false;
                }
            }

            value = value.Replace("$", string.Empty);
            if (A1.TryParseRange(value, out _, out _, out _, out _)) {
                dataRange = value;
                return true;
            }

            if (A1.TryParseCellReferenceFast(value, out _, out _)) {
                dataRange = value;
                return true;
            }

            return false;
        }

        private static string UnquoteSheetName(string value) {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length >= 2 && normalized[0] == '\'' && normalized[normalized.Length - 1] == '\'') {
                normalized = normalized.Substring(1, normalized.Length - 2).Replace("''", "'");
            }

            return normalized;
        }

        private static bool TryReadSparklineValues(
            ExcelSheet sheet,
            string dataRange,
            ref long remainingDataCells,
            out IReadOnlyList<double> values) {
            int firstRow;
            int firstColumn;
            int lastRow;
            int lastColumn;
            if (!A1.TryParseRange(dataRange, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                if (!A1.TryParseCellReferenceFast(dataRange, out firstRow, out firstColumn)) {
                    values = Array.Empty<double>();
                    return false;
                }

                lastRow = firstRow;
                lastColumn = firstColumn;
            }

            long rowCount = (long)lastRow - firstRow + 1L;
            long columnCount = (long)lastColumn - firstColumn + 1L;
            if (rowCount <= 0L || columnCount <= 0L || rowCount > remainingDataCells / columnCount) {
                values = Array.Empty<double>();
                return false;
            }

            remainingDataCells -= rowCount * columnCount;

            var resolvedValues = new List<double>();
            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    ExcelCellData data = sheet.GetCellValueSnapshot(row, column);
                    if (TryConvertSparklineNumber(data.Value, out double value) ||
                        TryConvertSparklineNumber(data.CachedText, out value)) {
                        resolvedValues.Add(value);
                    }
                }
            }

            values = resolvedValues.AsReadOnly();
            return true;
        }

        private static bool TryConvertSparklineNumber(object? value, out double number) {
            if (value is double doubleValue) {
                number = doubleValue;
                return true;
            }

            if (value is IConvertible convertible &&
                double.TryParse(Convert.ToString(convertible, CultureInfo.InvariantCulture), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number)) {
                return true;
            }

            number = 0D;
            return false;
        }

        private static void ReportUnsupportedDrawingObject(
            ExcelSheet sheet,
            ExcelWorksheetDrawingObjectInfo drawing,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            List<OfficeImageExportDiagnostic> diagnostics) {
            string? sourceCell = drawing.CellReference;
            if (sourceCell != null
                && !TryResolveVisibleExportCell(sourceCell, firstRow, firstColumn, lastRow, lastColumn, columnsByIndex, rowsByIndex, out sourceCell)) {
                return;
            }

            string reason = string.IsNullOrWhiteSpace(drawing.UnsupportedReason)
                ? "not rendered by the dependency-free image exporter yet"
                : drawing.UnsupportedReason!;
            string source = sourceCell == null
                ? sheet.Name + "!" + drawing.Name
                : sheet.Name + "!" + sourceCell;
            diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported,
                "Worksheet drawing object '" + drawing.Name + "' (" + drawing.Kind + ") is " + reason + ".",
                source));
        }

        private static bool TryResolveVisibleExportCell(
            string reference,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            out string cellReference) {
            cellReference = string.Empty;
            string normalized = reference.Trim().Replace("$", string.Empty);
            (int row, int column) = A1.ParseCellRef(normalized);
            if (row < firstRow || row > lastRow || column < firstColumn || column > lastColumn) {
                return false;
            }

            if (!rowsByIndex.ContainsKey(row) || !columnsByIndex.ContainsKey(column)) {
                return false;
            }

            cellReference = A1.CellReference(row, column);
            return true;
        }

        private static bool TryGetImageAnchor(
            ExcelImageExportOptions options,
            int firstRow,
            int firstColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            int columnIndex,
            int rowIndex,
            int offsetXPixels,
            int offsetYPixels,
            int widthPixels,
            int heightPixels,
            int? toColumnIndex,
            int? toRowIndex,
            int toOffsetXPixels,
            int toOffsetYPixels,
            out double x,
            out double y,
            out double width,
            out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (rowIndex <= 0 || columnIndex <= 0 || columnsByIndex.Count == 0 || rowsByIndex.Count == 0) {
                return false;
            }

            int lastVisibleColumn = columnsByIndex.Keys.Max();
            int lastVisibleRow = rowsByIndex.Keys.Max();
            if (DistanceOutsideRange(columnIndex, firstColumn, lastVisibleColumn) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                DistanceOutsideRange(rowIndex, firstRow, lastVisibleRow) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                (toColumnIndex.HasValue &&
                    (DistanceOutsideRange(toColumnIndex.Value, firstColumn, lastVisibleColumn) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                     Math.Abs((long)toColumnIndex.Value - columnIndex) > ExcelImageExportLimits.MaximumAnchorSpanCells)) ||
                (toRowIndex.HasValue &&
                    (DistanceOutsideRange(toRowIndex.Value, firstRow, lastVisibleRow) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                     Math.Abs((long)toRowIndex.Value - rowIndex) > ExcelImageExportLimits.MaximumAnchorSpanCells))) {
                return false;
            }

            offsetXPixels = ExcelImageExportLimits.ClampOffsetPixels(offsetXPixels);
            offsetYPixels = ExcelImageExportLimits.ClampOffsetPixels(offsetYPixels);
            toOffsetXPixels = ExcelImageExportLimits.ClampOffsetPixels(toOffsetXPixels);
            toOffsetYPixels = ExcelImageExportLimits.ClampOffsetPixels(toOffsetYPixels);
            widthPixels = ExcelImageExportLimits.ClampExtentPixels(widthPixels);
            heightPixels = ExcelImageExportLimits.ClampExtentPixels(heightPixels);
            x = ResolveRelativeColumnOffset(firstColumn, columnIndex, offsetXPixels, columnDefinitions, options);
            y = ResolveRelativeRowOffset(firstRow, rowIndex, offsetYPixels, rowDefinitions, defaultRowsHidden, options);
            width = ResolveImageWidth(options, firstColumn, columnDefinitions, columnIndex, offsetXPixels, widthPixels, toColumnIndex, toOffsetXPixels);
            height = ResolveImageHeight(options, firstRow, rowDefinitions, defaultRowsHidden, rowIndex, offsetYPixels, heightPixels, toRowIndex, toOffsetYPixels);

            double rangeWidth = ResolveVisualWidth(columnsByIndex);
            double rangeHeight = ResolveVisualHeight(rowsByIndex);
            return x < rangeWidth &&
                x + width > 0D &&
                y < rangeHeight &&
                y + height > 0D;
        }

        private static long DistanceOutsideRange(int value, int first, int last) {
            if (value < first) {
                return (long)first - value;
            }

            return value > last ? (long)value - last : 0L;
        }

        private static bool TryGetAnchor(
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            int columnIndex,
            int rowIndex,
            int widthPixels,
            int heightPixels,
            out double x,
            out double y,
            out double width,
            out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (!columnsByIndex.TryGetValue(columnIndex, out ExcelVisualColumn? column) ||
                !rowsByIndex.TryGetValue(rowIndex, out ExcelVisualRow? row)) {
                return false;
            }

            x = column.X;
            y = row.Y;
            width = Math.Max(1D, widthPixels);
            height = Math.Max(1D, heightPixels);
            return true;
        }

        private static double ResolveRelativeColumnOffset(
            int firstColumn,
            int anchorColumn,
            int offsetPixels,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            ExcelImageExportOptions options) {
            double x = 0D;
            if (anchorColumn >= firstColumn) {
                for (int column = firstColumn; column < anchorColumn; column++) {
                    x += ResolveVisibleColumnWidth(column, columnDefinitions, options);
                }
            } else {
                for (int column = anchorColumn; column < firstColumn; column++) {
                    x -= ResolveVisibleColumnWidth(column, columnDefinitions, options);
                }
            }

            return x + offsetPixels;
        }

        private static double ResolveImageWidth(
            ExcelImageExportOptions options,
            int firstColumn,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            int anchorColumn,
            int offsetPixels,
            int widthPixels,
            int? toColumnIndex,
            int toOffsetPixels) {
            if (toColumnIndex.HasValue) {
                double from = ResolveRelativeColumnOffset(firstColumn, anchorColumn, offsetPixels, columnDefinitions, options);
                double to = ResolveRelativeColumnOffset(firstColumn, toColumnIndex.Value, toOffsetPixels, columnDefinitions, options);
                if (to > from) {
                    return Math.Min(ExcelImageExportLimits.MaximumAnchorExtentPixels, Math.Max(1D, to - from));
                }
            }

            return Math.Max(1D, widthPixels);
        }

        private static double ResolveRelativeRowOffset(
            int firstRow,
            int anchorRow,
            int offsetPixels,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            ExcelImageExportOptions options) {
            double y = 0D;
            if (anchorRow >= firstRow) {
                for (int row = firstRow; row < anchorRow; row++) {
                    y += ResolveVisibleRowHeight(row, rowDefinitions, defaultRowsHidden, options);
                }
            } else {
                for (int row = anchorRow; row < firstRow; row++) {
                    y -= ResolveVisibleRowHeight(row, rowDefinitions, defaultRowsHidden, options);
                }
            }

            return y + offsetPixels;
        }

        private static double ResolveImageHeight(
            ExcelImageExportOptions options,
            int firstRow,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            int anchorRow,
            int offsetPixels,
            int heightPixels,
            int? toRowIndex,
            int toOffsetPixels) {
            if (toRowIndex.HasValue) {
                double from = ResolveRelativeRowOffset(firstRow, anchorRow, offsetPixels, rowDefinitions, defaultRowsHidden, options);
                double to = ResolveRelativeRowOffset(firstRow, toRowIndex.Value, toOffsetPixels, rowDefinitions, defaultRowsHidden, options);
                if (to > from) {
                    return Math.Min(ExcelImageExportLimits.MaximumAnchorExtentPixels, Math.Max(1D, to - from));
                }
            }

            return Math.Max(1D, heightPixels);
        }

        private static double ResolveVisibleColumnWidth(int column, IReadOnlyList<ExcelColumnSnapshot> columnDefinitions, ExcelImageExportOptions options) {
            ExcelColumnSnapshot? definition = columnDefinitions.FirstOrDefault(item => column >= item.StartIndex && column <= item.EndIndex);
            if (definition?.Hidden == true && !options.IncludeHidden) {
                return 0D;
            }

            return ResolveColumnWidth(definition, options);
        }

        private static double ResolveVisibleRowHeight(int row, IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions, bool defaultRowsHidden, ExcelImageExportOptions options) {
            rowDefinitions.TryGetValue(row, out ExcelRowSnapshot? definition);
            if (IsHiddenRow(row, rowDefinitions, defaultRowsHidden) && !options.IncludeHidden) {
                return 0D;
            }

            return ResolveRowHeight(definition, options);
        }

        private static double ResolveVisualWidth(IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex) =>
            columnsByIndex.Count == 0 ? 0D : columnsByIndex.Values.Max(column => column.X + column.Width);

        private static double ResolveVisualHeight(IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex) =>
            rowsByIndex.Count == 0 ? 0D : rowsByIndex.Values.Max(row => row.Y + row.Height);

        private static void ReportHiddenRowColumnOmissions(
            ExcelSheet sheet,
            string range,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            List<OfficeImageExportDiagnostic> diagnostics) {
            int hiddenRows = 0;
            for (int row = firstRow; row <= lastRow; row++) {
                if (IsHiddenRow(row, rowDefinitions, defaultRowsHidden)) {
                    hiddenRows++;
                }
            }

            int hiddenColumns = CountHiddenColumns(firstColumn, lastColumn, columnDefinitions);
            string source = sheet.Name + "!" + range;
            if (hiddenRows > 0) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.HiddenRowsOmitted,
                    hiddenRows.ToString(CultureInfo.InvariantCulture) + " hidden row(s) were omitted from the Excel image export.",
                    source));
            }

            if (hiddenColumns > 0) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.HiddenColumnsOmitted,
                    hiddenColumns.ToString(CultureInfo.InvariantCulture) + " hidden column(s) were omitted from the Excel image export.",
                    source));
            }
        }

        private static int CountHiddenColumns(int firstColumn, int lastColumn, IReadOnlyList<ExcelColumnSnapshot> columnDefinitions) {
            var hiddenColumns = new HashSet<int>();
            foreach (ExcelColumnSnapshot definition in columnDefinitions) {
                if (!definition.Hidden) {
                    continue;
                }

                int start = Math.Max(firstColumn, definition.StartIndex);
                int end = Math.Min(lastColumn, definition.EndIndex);
                for (int column = start; column <= end; column++) {
                    hiddenColumns.Add(column);
                }
            }

            return hiddenColumns.Count;
        }

        private static bool IsHiddenAnchorInRange(
            int rowIndex,
            int columnIndex,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions) {
            return rowIndex >= firstRow &&
                rowIndex <= lastRow &&
                columnIndex >= firstColumn &&
                columnIndex <= lastColumn &&
                (IsHiddenRow(rowIndex, rowDefinitions, defaultRowsHidden) || IsHiddenColumn(columnIndex, columnDefinitions));
        }

        private static bool IsHiddenRow(int rowIndex, IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions, bool defaultRowsHidden) {
            if (rowDefinitions.TryGetValue(rowIndex, out ExcelRowSnapshot? definition)) {
                return definition.Hidden;
            }

            return defaultRowsHidden;
        }

        private static bool IsHiddenColumn(int columnIndex, IReadOnlyList<ExcelColumnSnapshot> columnDefinitions) {
            return columnDefinitions.Any(definition =>
                definition.Hidden &&
                columnIndex >= definition.StartIndex &&
                columnIndex <= definition.EndIndex);
        }

        private static bool Intersects(int r1, int c1, int r2, int c2, int otherR1, int otherC1, int otherR2, int otherC2) =>
            r1 <= otherR2 && r2 >= otherR1 && c1 <= otherC2 && c2 >= otherC1;

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

        private static string FormatCellDisplayText(string text, ExcelCellStyleSnapshot style, ExcelDateSystem dateSystem) {
            if (string.IsNullOrEmpty(text) || style.NumberFormatId == 0U) {
                return text;
            }

            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? ExcelNumberFormatDisplay.FormatNumericText(value, style.NumberFormatId, style.NumberFormatCode, text, dateSystem)
                : text;
        }

        private static ExcelVisualCellValueKind ResolveVisualCellValueKind(ExcelCellData data, string rawText, ExcelCellStyleSnapshot style) {
            switch (data.Kind) {
                case ExcelCellDataKind.Blank:
                    return ExcelVisualCellValueKind.Blank;
                case ExcelCellDataKind.Boolean:
                    return ExcelVisualCellValueKind.Boolean;
                case ExcelCellDataKind.Error:
                    return ExcelVisualCellValueKind.Error;
                case ExcelCellDataKind.Text:
                    return ExcelVisualCellValueKind.Text;
                case ExcelCellDataKind.Number:
                    return IsDateLikeNumericCell(style) ? ExcelVisualCellValueKind.Date : ExcelVisualCellValueKind.Number;
                case ExcelCellDataKind.Formula:
                    return TryResolveFormulaVisualValueKind(data, rawText, style, out ExcelVisualCellValueKind formulaKind)
                        ? formulaKind
                        : ExcelVisualCellValueKind.Text;
                default:
                    return ExcelVisualCellValueKind.Text;
            }
        }

        private static bool TryResolveFormulaVisualValueKind(ExcelCellData data, string rawText, ExcelCellStyleSnapshot style, out ExcelVisualCellValueKind valueKind) {
            string? cachedText = string.IsNullOrEmpty(data.CachedText) ? rawText : data.CachedText;
            if (data.Value is double || double.TryParse(cachedText, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) {
                valueKind = IsDateLikeNumericCell(style) ? ExcelVisualCellValueKind.Date : ExcelVisualCellValueKind.Number;
                return true;
            }

            valueKind = ExcelVisualCellValueKind.Text;
            return false;
        }

        private static bool IsDateLikeNumericCell(ExcelCellStyleSnapshot style) =>
            style.IsDateLike || ExcelNumberFormatDisplay.IsDateNumberFormat(style.NumberFormatId, style.NumberFormatCode);

        private static string GetImageDiagnosticSource(ExcelSheet sheet, ExcelImage image) {
            string name = string.IsNullOrWhiteSpace(image.Name) ? "Image" : image.Name;
            return sheet.Name + "!" + name;
        }

        private static string GetDrawingDiagnosticSource(ExcelSheet sheet, ExcelWorksheetDrawingObjectInfo drawing) {
            if (!string.IsNullOrWhiteSpace(drawing.CellReference)) {
                return sheet.Name + "!" + drawing.CellReference;
            }

            string name = string.IsNullOrWhiteSpace(drawing.Name) ? "Drawing" : drawing.Name;
            return sheet.Name + "!" + name;
        }

        private static string Key(int row, int column) => row.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + column.ToString(System.Globalization.CultureInfo.InvariantCulture);

        private static string MergeGeometryKey(ExcelMergedRangeSnapshot merge) =>
            Key(merge.StartRow, merge.StartColumn) + ":" + Key(merge.EndRow, merge.EndColumn);
    }
}
