using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private const double WorksheetHeadingHeight = 22D;

        private static void AddWorksheetCanvasPages(
            PdfCore.PdfDocument pdf,
            ExcelDocument document,
            IReadOnlyList<WorksheetPdfExportPlan> plans,
            ExcelToPdfOptions options,
            IReadOnlyDictionary<string, string> sheetDestinations,
            IReadOnlyDictionary<string, string> cellDestinations,
            PdfCore.PdfStandardFont defaultFontFamily) {
            WorksheetPdfExportPlan firstPlan = plans[0];
            PdfCore.PageSize pageSize = GetEffectivePageSize(options, firstPlan.PageSetup);
            PdfCore.PageMargins margins = GetEffectiveMargins(options, firstPlan.PageSetup);
            double headingHeight = options.IncludeSheetHeadings ? WorksheetHeadingHeight : 0D;
            double availableWidth = Math.Max(1D, pageSize.Width - margins.Left - margins.Right);
            double availableHeight = Math.Max(1D, pageSize.Height - margins.Top - margins.Bottom - headingHeight);
            pdf.Section(page => {
                ApplyWorksheetPageSetup(page, firstPlan.PageSetup, options);
                ApplyWorksheetHeaderFooter(page, firstPlan.HeaderFooter, firstPlan.SheetName, document.FilePath, options);
                for (int areaIndex = 0; areaIndex < plans.Count; areaIndex++) {
                    WorksheetPdfExportPlan plan = plans[areaIndex];
                    int columns = plan.ExportData.Values.GetLength(1);
                    double worksheetScale = ResolveWorksheetPlanScale(plan, columns, availableWidth, availableHeight);
                    IReadOnlyList<TableChunk> chunks = plan.HasTable
                        ? CreateWorksheetSceneChunks(plan, options, columns, availableWidth, availableHeight, worksheetScale)
                        : new[] { new TableChunk(Array.Empty<int>(), 0, 0, 0) };
                    // Vertical pages share the same column projection. Keep one per horizontal
                    // segment for this print area instead of copying all source rows on every page.
                    if (areaIndex > 0) page.Content(content => content.Item(item => item.PageBreak()));
                    page.Content(content => content.Item(item => item.Bookmark(plan.BookmarkName)));
                    for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                        if (chunkIndex > 0) {
                            page.Content(content => content.Item(item => item.PageBreak()));
                        }

                        TableChunk chunk = chunks[chunkIndex];
                        bool firstPageForSheet = areaIndex == 0 && chunkIndex == 0;
                        page.Canvas(canvas => RenderWorksheetScene(
                            canvas,
                            plan,
                            chunk,
                            options,
                            sheetDestinations,
                            cellDestinations,
                            defaultFontFamily,
                            margins,
                            availableWidth,
                            availableHeight,
                            worksheetScale,
                            headingHeight,
                            firstPageForSheet,
                            document.DateSystem));
                    }
                }
            });
        }

        private static void RenderWorksheetScene(
            PdfCore.PdfPageCanvas canvas,
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            ExcelToPdfOptions options,
            IReadOnlyDictionary<string, string> sheetDestinations,
            IReadOnlyDictionary<string, string> cellDestinations,
            PdfCore.PdfStandardFont defaultFontFamily,
            PdfCore.PageMargins margins,
            double availableWidth,
            double availableHeight,
            double worksheetScale,
            double headingHeight,
            bool firstPageForSheet,
            ExcelDateSystem dateSystem) {
            double sceneX = margins.Left;
            double sceneY = margins.Top + headingHeight;
            if (firstPageForSheet) {
                canvas.Outline(plan.SheetName, 1, margins.Top);
            }

            if (headingHeight > 0D) {
                canvas.Text(
                    new[] { PdfCore.PdfTextRun.Bolded(plan.SheetName, fontSize: 15D) },
                    PdfCore.PdfCanvasTextStructureRole.Heading1,
                    margins.Left,
                    margins.Top,
                    availableWidth,
                    headingHeight,
                    fontSize: 15D);
            }

            List<double> columnWidths = CreateWorksheetSceneColumnWidths(plan, chunk);
            List<double> rowHeights = CreateWorksheetSceneRowHeights(plan, chunk);
            double tableWidth = Math.Max(1D, columnWidths.Sum());
            double tableHeight = Math.Max(1D, rowHeights.Sum());
            WorksheetSceneBounds objectBounds = MeasureWorksheetObjects(plan, chunk, columnWidths, rowHeights);
            // Media anchored outside a print/export range is filtered while the plan is built.
            // Retained objects therefore belong to this scene and must contribute their full
            // authored extent; clipping them to the final cell boundary can leave only a chart
            // title or a thin image strip visible when the anchor is in the last exported row.
            double unscaledSceneWidth = Math.Max(tableWidth, objectBounds.Right) - objectBounds.Left;
            double unscaledSceneHeight = Math.Max(tableHeight, objectBounds.Bottom) - objectBounds.Top;
            double scale = ResolveWorksheetSceneScale(plan.PageSetup, unscaledSceneWidth, unscaledSceneHeight, availableWidth, availableHeight, worksheetScale);
            sceneX += ExcelPageSetupGeometry.CenteredOffset(plan.PrintOptions?.HorizontalCentered, availableWidth, unscaledSceneWidth * scale);
            sceneY += ExcelPageSetupGeometry.CenteredOffset(plan.PrintOptions?.VerticalCentered, availableHeight, unscaledSceneHeight * scale);
            double clipWidth = Math.Min(availableWidth, unscaledSceneWidth * scale);
            double clipHeight = Math.Min(availableHeight, unscaledSceneHeight * scale);

            double tableX = sceneX - objectBounds.Left * scale;
            double tableY = sceneY - objectBounds.Top * scale;
            canvas.Clip(sceneX, sceneY, Math.Max(1D, clipWidth), Math.Max(1D, clipHeight), clipped => {
                if (plan.HasTable && chunk.RowIndexes.Count > 0 && chunk.ColumnCount > 0) {
                    SheetExportData data = SelectPageCells(plan, chunk, options);
                    int[] pageRows = Enumerable.Range(0, chunk.RowIndexes.Count).ToArray();
                    PdfCore.PdfTableStyle tableStyle = CreateWorksheetSceneTableStyle(plan, data, chunk, options, columnWidths, rowHeights, scale);
                    clipped.Table(
                        CreatePdfRows(
                            data.Values,
                            data.Styles,
                            data.Hyperlinks,
                            data.CellReferences,
                            data.StructuredTables,
                            data.MergedCells,
                            imagesByCellReference: null,
                            pageRows,
                            0,
                            chunk.ColumnIndexes.Count,
                            options.EmptyCellText,
                            sheetDestinations,
                            cellDestinations,
                            plan.SheetName,
                            defaultFontFamily,
                            scale,
                            preserveWorksheetNoWrap: true,
                            dateSystem: dateSystem),
                        tableX,
                        tableY,
                        tableWidth * scale,
                        tableHeight * scale,
                        tableStyle);
                    AddWorksheetSceneConditionalIcons(clipped, data, columnWidths, rowHeights, tableX, tableY, scale);
                }

                AddWorksheetSceneImages(clipped, plan, chunk, columnWidths, rowHeights, tableX, tableY, scale);
                AddWorksheetSceneCharts(clipped, plan, chunk, options, columnWidths, rowHeights, tableX, tableY, scale);
                AddWorksheetSceneDrawingObjects(clipped, plan, chunk, columnWidths, rowHeights, tableX, tableY, scale);
            });
        }

        private static PdfCore.PdfTableStyle CreateWorksheetSceneTableStyle(
            WorksheetPdfExportPlan plan,
            SheetExportData data,
            TableChunk chunk,
            ExcelToPdfOptions options,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            double scale) {
            PdfCore.PdfTableStyle style = CreateTableStyle(
                options,
                plan.PageSetup,
                Enumerable.Range(0, chunk.RowIndexes.Count).ToArray(),
                chunk.HeaderRowCount,
                data.Styles,
                data.ConditionalFills,
                data.CellReferences,
                data.StructuredTables,
                data.ColumnWidths,
                data.RowHeights,
                0,
                chunk.ColumnIndexes.Count);
            style.HeaderFill = null;
            ApplyGeneralCellAlignments(style, data.Values, Enumerable.Range(0, chunk.RowIndexes.Count).ToArray());
            style.RightAlignNumeric = false;
            style.HeaderBold = false;
            style.HeaderTextColor = null;
            style.FooterFill = null;
            style.RowStripeFill = null;
            style.BorderColor = null;
            style.ColumnWidthPoints = columnWidths.Select(width => (double?)width).ToList();
            style.FixedRowHeights = rowHeights.Select(height => (double?)height).ToList();
            style.RowMinHeights = null;
            style.CellPaddingX = 1.5D * scale;
            style.CellPaddingY = 0D;
            style.CellPaddingLeft = null;
            style.CellPaddingRight = null;
            style.CellPaddingTop = null;
            style.CellPaddingBottom = null;
            style.FontSize = GetDefaultTableFontSize(options) * scale;
            style.HeaderFontSize = style.FontSize;
            style.FooterFontSize = style.FontSize;
            style.LineHeight = 1D;
            style.VerticalAlignments = Enumerable.Repeat(PdfCore.PdfCellVerticalAlign.Bottom, chunk.ColumnIndexes.Count).ToList();
            style.MinRowHeight = 0D;
            style.MaxWidth = null;
            style.PreserveWidth = true;
            style.ShrinkTextToFit = false;
            style.MinimumShrinkFontSize = Math.Max(3D, 4D * scale);
            style.SpacingBefore = 0D;
            style.SpacingAfter = 0D;
            style.CellIcons = null;
            ApplyWorksheetConditionalIconStyle(style, data, columnWidths, rowHeights, scale);
            style.AlternativeText = "Worksheet " + plan.SheetName;
            return style;
        }

        private static List<double> CreateWorksheetSceneColumnWidths(WorksheetPdfExportPlan plan, TableChunk chunk) {
            var widths = new List<double>(chunk.ColumnIndexes.Count);
            foreach (int column in chunk.ColumnIndexes) {
                widths.Add(GetExportedColumnWidthPoints(plan, column));
            }

            return widths;
        }

        private static List<double> CreateWorksheetSceneRowHeights(WorksheetPdfExportPlan plan, TableChunk chunk) =>
            chunk.RowIndexes.Select(row => GetExportedRowHeightPoints(plan, row)).ToList();

        private static double GetExportedColumnWidthPoints(WorksheetPdfExportPlan plan, int exportedColumn) {
            string?[,]? references = plan.ExportData.CellReferences;
            int rows = references == null ? 0 : Math.Min(plan.ExportedRows, references.GetLength(0));
            int sourceColumn = references == null ? 0 : GetOriginalColumnNumber(references, exportedColumn, rows);
            if (sourceColumn <= 0) {
                sourceColumn = plan.Geometry.FirstColumn + exportedColumn;
            }

            return GetWorksheetColumnWidthPoints(plan.Geometry, sourceColumn);
        }

        private static double GetExportedRowHeightPoints(WorksheetPdfExportPlan plan, int exportedRow) {
            string?[,]? references = plan.ExportData.CellReferences;
            int sourceRow = references == null ? 0 : GetOriginalRowNumber(references, exportedRow);
            if (sourceRow <= 0) {
                sourceRow = plan.Geometry.FirstRow + exportedRow;
            }

            return GetWorksheetRowHeightPoints(plan.Geometry, sourceRow);
        }

        private static double GetWorksheetColumnWidthPoints(WorksheetGeometryData geometry, int sourceColumn) {
            double width = geometry.DefaultColumnWidth;
            bool hasExplicitWidth = geometry.HasExplicitDefaultColumnWidth;
            for (int index = geometry.Columns.Count - 1; index >= 0; index--) {
                ExcelColumnSnapshot definition = geometry.Columns[index];
                if (sourceColumn < definition.StartIndex || sourceColumn > definition.EndIndex) {
                    continue;
                }

                if (geometry.RespectHiddenRowsAndColumns && definition.Hidden) {
                    return 0D;
                }

                if (geometry.UseWorksheetColumnWidths &&
                    definition.CustomWidth &&
                    definition.Width is double customWidth &&
                    customWidth > 0D) {
                    width = customWidth;
                    hasExplicitWidth = true;
                }
                break;
            }

            if (geometry.MaximumDigitWidthPoints is double maximumDigitWidth) {
                if (!hasExplicitWidth) {
                    // baseColWidth excludes the margin and gridline allowance.
                    // The five logical pixels are format padding at 96 DPI.
                    const double paddingPoints = 5D * 72D / 96D;
                    return Math.Max(0.75D, geometry.BaseColumnWidth * maximumDigitWidth + paddingPoints);
                }
                // The stored width already includes the character-unit padding.
                // Keep PDF geometry in points rather than quantizing it to a screen
                // pixel grid. Printer-specific hinting remains a separate limit.
                return Math.Max(0.75D, width * maximumDigitWidth);
            }

            // Preserve the established estimate when the workbook default font
            // program is unavailable; a substituted text font does not establish
            // the workbook's column units.
            const double fallbackDigitWidth = 7D;
            double pixels = Math.Truncate((256D * width + Math.Truncate(128D / fallbackDigitWidth)) / 256D * fallbackDigitWidth) + 5D;
            return Math.Max(0.75D, Math.Max(1D, pixels) * 72D / 96D);
        }

        private static double GetWorksheetRowHeightPoints(WorksheetGeometryData geometry, int sourceRow) {
            for (int index = geometry.Rows.Count - 1; index >= 0; index--) {
                ExcelRowSnapshot definition = geometry.Rows[index];
                if (definition.Index != sourceRow) {
                    continue;
                }
                return GetWorksheetDefinedRowHeightPoints(geometry, definition);
            }

            return geometry.DefaultRowHeight;
        }

        private static double GetWorksheetDefinedRowHeightPoints(WorksheetGeometryData geometry, ExcelRowSnapshot definition) =>
            geometry.RespectHiddenRowsAndColumns && definition.Hidden ? 0D
                : geometry.UseWorksheetRowHeights && definition.CustomHeight &&
                  definition.Height is double customHeight && customHeight > 0D
                    ? customHeight : geometry.DefaultRowHeight;

        private static double GetWorksheetRowHeightRangePoints(WorksheetGeometryData geometry, int firstRow, int endRow) {
            if (endRow <= firstRow) return 0D;
            double total = (endRow - firstRow) * geometry.DefaultRowHeight;
            var seen = new HashSet<int>();
            for (int index = geometry.Rows.Count - 1; index >= 0; index--) {
                ExcelRowSnapshot definition = geometry.Rows[index];
                if (definition.Index < firstRow || definition.Index >= endRow || !seen.Add(definition.Index)) continue;
                total += GetWorksheetDefinedRowHeightPoints(geometry, definition) - geometry.DefaultRowHeight;
            }
            return total;
        }

        private static double ResolveWorksheetSceneScale(
            ExcelSheetPageSetup? pageSetup,
            double sceneWidth,
            double sceneHeight,
            double availableWidth,
            double availableHeight,
            double worksheetScale) {
            double scale = worksheetScale;
            if (IsFitToWidth(pageSetup) || sceneWidth * scale > availableWidth) {
                scale = Math.Min(scale, availableWidth / Math.Max(1D, sceneWidth));
            }

            if (IsFitToHeight(pageSetup) || sceneHeight * scale > availableHeight) {
                scale = Math.Min(scale, availableHeight / Math.Max(1D, sceneHeight));
            }

            // A lower bound here can push retained cells outside the page when media
            // precedes a distant used range. Both computed ratios are already positive.
            return scale;
        }

        private static double GetWorksheetAuthoredScale(ExcelSheetPageSetup? pageSetup) {
            if (!ExcelPageSetupGeometry.HasFitToPageScale(pageSetup) &&
                pageSetup?.Scale is uint scale &&
                scale > 0U) {
                return Math.Max(0.1D, Math.Min(4D, scale / 100D));
            }

            return 1D;
        }

        private static bool IsFitToWidth(ExcelSheetPageSetup? pageSetup) =>
            ExcelPageSetupGeometry.HasFitToPageScale(pageSetup) &&
            pageSetup?.FitToWidth is uint fitToWidth &&
            fitToWidth > 0U;

        private static WorksheetSceneBounds MeasureWorksheetObjects(
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights) {
            double left = 0D;
            double top = 0D;
            double right = 0D;
            double bottom = 0D;
            foreach (WorksheetImageExportData image in plan.Images) {
                if (!TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, image.RowIndex, image.ColumnIndex, out double x, out double y)) {
                    continue;
                }

                left = Math.Min(left, x + image.OffsetXPoints);
                top = Math.Min(top, y + image.OffsetYPoints);
                right = Math.Max(right, x + image.OffsetXPoints + image.WidthPoints);
                bottom = Math.Max(bottom, y + image.OffsetYPoints + image.HeightPoints);
            }

            foreach (WorksheetChartExportData chart in plan.Charts) {
                ExcelChartSnapshot snapshot = chart.Snapshot;
                if (!TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, snapshot.RowIndex, snapshot.ColumnIndex, out double x, out double y)) {
                    continue;
                }

                left = Math.Min(left, x + PixelsToPoints(snapshot.OffsetXPixels));
                top = Math.Min(top, y + PixelsToPoints(snapshot.OffsetYPixels));
                right = Math.Max(right, x + PixelsToPoints(snapshot.OffsetXPixels + snapshot.WidthPixels));
                bottom = Math.Max(bottom, y + PixelsToPoints(snapshot.OffsetYPixels + snapshot.HeightPixels));
            }

            foreach (var drawing in plan.DrawingObjects) {
                if (!TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, drawing.Row, drawing.Column, out double x, out double y)) continue;
                (double width, double height) = GetWorksheetDrawingSize(plan, drawing);
                left = Math.Min(left, x + PixelsToPoints(drawing.OffsetXPixels));
                top = Math.Min(top, y + PixelsToPoints(drawing.OffsetYPixels));
                right = Math.Max(right, x + PixelsToPoints(drawing.OffsetXPixels) + width);
                bottom = Math.Max(bottom, y + PixelsToPoints(drawing.OffsetYPixels) + height);
            }
            return new WorksheetSceneBounds(left, top, right, bottom);
        }

        private static void AddWorksheetSceneImages(
            PdfCore.PdfPageCanvas canvas,
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            double sceneX,
            double sceneY,
            double scale) {
            foreach (WorksheetImageExportData image in plan.Images) {
                if (!TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, image.RowIndex, image.ColumnIndex, out double x, out double y)) {
                    continue;
                }

                canvas.Image(
                    image.Bytes,
                    sceneX + (x + image.OffsetXPoints) * scale,
                    sceneY + (y + image.OffsetYPoints) * scale,
                    image.WidthPoints * scale,
                    image.HeightPoints * scale,
                    CreateConverterImageStyle(),
                    alternativeText: image.AlternativeText,
                    rotationAngle: -image.RotationDegrees,
                    horizontalFlip: image.HorizontalFlip,
                    verticalFlip: image.VerticalFlip);
            }
        }

        private static void AddWorksheetSceneCharts(
            PdfCore.PdfPageCanvas canvas,
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            ExcelToPdfOptions options,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            double sceneX,
            double sceneY,
            double scale) {
            foreach (WorksheetChartExportData chart in plan.Charts) {
                ExcelChartSnapshot snapshot = chart.Snapshot;
                if (!TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, snapshot.RowIndex, snapshot.ColumnIndex, out double x, out double y)) {
                    continue;
                }

                OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(CreateOfficeChartSnapshotCore(snapshot, options, preserveWorksheetLegend: true));
                AddChartQualityWarning(options, plan.SheetName, snapshot, rendering.QualityReport);
                string chartName = string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title!;
                canvas.Drawing(
                    rendering.Drawing,
                    sceneX + (x + PixelsToPoints(snapshot.OffsetXPixels)) * scale,
                    sceneY + (y + PixelsToPoints(snapshot.OffsetYPixels)) * scale,
                    PixelsToPoints(snapshot.WidthPixels) * scale,
                    PixelsToPoints(snapshot.HeightPixels) * scale,
                    new PdfCore.PdfDrawingStyle {
                        AlternativeText = string.IsNullOrWhiteSpace(chartName) ? "Worksheet chart" : chartName
                    });
            }
        }

        private static bool TryGetWorksheetObjectOffset(
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            IReadOnlyList<double> columnWidths,
            IReadOnlyList<double> rowHeights,
            int sourceRow,
            int sourceColumn,
            out double x,
            out double y) {
            x = 0D;
            y = 0D;
            string?[,]? references = plan.ExportData.CellReferences;
            if (references == null || chunk.RowIndexes.Count == 0 || chunk.ColumnCount <= 0) {
                if (plan.HasTable ||
                    sourceRow < plan.Geometry.FirstRow ||
                    sourceColumn < plan.Geometry.FirstColumn) {
                    return false;
                }

                y += GetWorksheetRowHeightRangePoints(plan.Geometry, plan.Geometry.FirstRow, sourceRow);
                for (int column = plan.Geometry.FirstColumn; column < sourceColumn; column++) {
                    x += GetWorksheetColumnWidthPoints(plan.Geometry, column);
                }
                return true;
            }

            int localRow = -1;
            for (int index = 0; index < chunk.RowIndexes.Count; index++) {
                if (GetOriginalRowNumber(references, chunk.RowIndexes[index]) == sourceRow) {
                    localRow = index;
                    break;
                }
            }

            if (localRow >= 0 &&
                chunk.RowIndexes[localRow] < plan.ExportData.HeaderRowCount &&
                !IsInitialWorksheetRowChunk(plan, chunk)) {
                return false;
            }

            if (localRow < 0) {
                if (plan.ClipToExportRange) {
                    return false;
                }

                int lastExportedRow = Math.Min(plan.ExportedRows, references.GetLength(0)) - 1;
                int lastSourceRow = GetLastOriginalRowNumber(references, lastExportedRow);
                if (sourceRow < plan.Geometry.FirstRow && IsInitialWorksheetRowChunk(plan, chunk)) {
                    // Virtual leading space belongs to canvas geometry, not a dense cell range.
                    y -= GetWorksheetRowHeightRangePoints(plan.Geometry, sourceRow, plan.Geometry.FirstRow);
                } else {
                    if (lastSourceRow <= 0 || sourceRow <= lastSourceRow || !chunk.RowIndexes.Contains(lastExportedRow)) return false;
                    for (int index = 0; index < rowHeights.Count; index++) y += rowHeights[index];
                    y += GetWorksheetRowHeightRangePoints(plan.Geometry, lastSourceRow + 1, sourceRow);
                }
            } else {
                for (int index = 0; index < localRow; index++) {
                    y += rowHeights[index];
                }
            }

            int localColumn = -1;
            for (int index = 0; index < chunk.ColumnIndexes.Count; index++) {
                int exportedColumn = chunk.ColumnIndexes[index];
                if (GetOriginalColumnNumber(references, exportedColumn, Math.Min(plan.ExportedRows, references.GetLength(0))) == sourceColumn) {
                    localColumn = index;
                    break;
                }
            }

            if (localColumn >= 0) {
                for (int index = 0; index < localColumn; index++) {
                    x += columnWidths[index];
                }
                return true;
            }

            if (!plan.ClipToExportRange && sourceColumn < plan.Geometry.FirstColumn && chunk.StartColumn == 0) {
                for (int column = sourceColumn; column < plan.Geometry.FirstColumn; column++)
                    x -= GetWorksheetColumnWidthPoints(plan.Geometry, column);
                return true;
            }

            int lastExportedColumn = references.GetLength(1) - 1;
            int lastSourceColumn = GetLastOriginalColumnNumber(
                references,
                lastExportedColumn,
                Math.Min(plan.ExportedRows, references.GetLength(0)));
            if (plan.ClipToExportRange ||
                lastSourceColumn <= 0 ||
                sourceColumn <= lastSourceColumn ||
                chunk.StartColumn + chunk.ColumnCount - 1 != lastExportedColumn) {
                return false;
            }

            // A final page can display repeated titles followed by a distant body segment.
            // Omitted source columns do not occupy space on that page; only the genuine
            // trailing gap beyond the last exported column contributes additional width.
            x = columnWidths.Sum();
            for (int column = lastSourceColumn + 1; column < sourceColumn; column++) {
                x += GetWorksheetColumnWidthPoints(plan.Geometry, column);
            }

            return true;
        }

        private static bool IsInitialWorksheetRowChunk(WorksheetPdfExportPlan plan, TableChunk chunk) {
            int headerRows = Math.Min(plan.ExportData.HeaderRowCount, plan.ExportedRows);
            return plan.ExportedRows <= headerRows || chunk.RowIndexes.Contains(headerRows);
        }

        private static int GetLastOriginalRowNumber(string?[,] references, int startRow) {
            for (int row = Math.Min(startRow, references.GetLength(0) - 1); row >= 0; row--) {
                int sourceRow = GetOriginalRowNumber(references, row);
                if (sourceRow > 0) {
                    return sourceRow;
                }
            }

            return 0;
        }

        private static int GetLastOriginalColumnNumber(string?[,] references, int startColumn, int rows) {
            for (int column = Math.Min(startColumn, references.GetLength(1) - 1); column >= 0; column--) {
                int sourceColumn = GetOriginalColumnNumber(references, column, rows);
                if (sourceColumn > 0) {
                    return sourceColumn;
                }
            }

            return 0;
        }

        private readonly struct WorksheetSceneBounds {
            public WorksheetSceneBounds(double left, double top, double right, double bottom) {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public double Left { get; }
            public double Top { get; }
            public double Right { get; }
            public double Bottom { get; }
        }
    }
}
