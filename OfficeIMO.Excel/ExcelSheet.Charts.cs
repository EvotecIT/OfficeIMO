using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Enumerates charts on the worksheet.
        /// </summary>
        public IEnumerable<ExcelChart> Charts {
            get {
                var drawingPart = _worksheetPart.DrawingsPart;
                if (drawingPart?.WorksheetDrawing == null) return Enumerable.Empty<ExcelChart>();
                return drawingPart.WorksheetDrawing
                    .Descendants<Xdr.GraphicFrame>()
                    .Where(frame => frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null)
                    .Select(frame => new ExcelChart(frame, drawingPart, this))
                    .ToList();
            }
        }

        /// <summary>
        /// Returns a chart by name (non-visual name), or null if not found.
        /// </summary>
        public ExcelChart? GetChart(string name) {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return Charts.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Writes chart data into the worksheet and returns the corresponding data range.
        /// </summary>
        public ExcelChartDataRange WriteChartData(ExcelChartData data, int startRow = 1, int startColumn = 1, string? categoryHeader = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (startRow <= 0 || startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));

            var cells = new List<(int Row, int Column, object Value)>();
            string header = categoryHeader ?? string.Empty;
            cells.Add((startRow, startColumn, header));

            for (int s = 0; s < data.Series.Count; s++) {
                cells.Add((startRow, startColumn + s + 1, data.Series[s].Name));
            }

            for (int i = 0; i < data.Categories.Count; i++) {
                int row = startRow + i + 1;
                cells.Add((row, startColumn, data.Categories[i]));
                for (int s = 0; s < data.Series.Count; s++) {
                    cells.Add((row, startColumn + s + 1, data.Series[s].Values[i]));
                }
            }

            CellValues(cells, null);

            return new ExcelChartDataRange(Name, startRow, startColumn, data.Categories.Count, data.Series.Count, hasHeaderRow: true);
        }

        /// <summary>
        /// Adds a chart to the worksheet using the provided data. Data is stored on a hidden chart data sheet.
        /// </summary>
        public ExcelChart AddChart(ExcelChartData data, int row, int column, int widthPixels = 640, int heightPixels = 360,
            ExcelChartType type = ExcelChartType.ColumnClustered, string? title = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (row <= 0 || column <= 0) throw new ArgumentOutOfRangeException(nameof(row));

            var dataSheet = _excelDocument.GetOrCreateChartDataSheet();
            int startRow = _excelDocument.ReserveChartDataStartRow(data.Categories.Count + 1);
            ExcelChartDataRange range = dataSheet.WriteChartData(data, startRow, 1);
            return AddChart(range, row, column, widthPixels, heightPixels, type, data, title);
        }

        /// <summary>
        /// Adds a chart to the worksheet using an existing data range.
        /// </summary>
        public ExcelChart AddChart(ExcelChartDataRange dataRange, int row, int column, int widthPixels = 640, int heightPixels = 360,
            ExcelChartType type = ExcelChartType.ColumnClustered, ExcelChartData? cachedData = null, string? title = null) {
            if (dataRange == null) throw new ArgumentNullException(nameof(dataRange));
            if (row <= 0 || column <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (widthPixels <= 0 || heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));

            return AddChartInternal(dataRange, row, column, widthPixels, heightPixels, type, cachedData, title);
        }

        /// <summary>
        /// Adds a chart using an A1 range on this sheet.
        /// </summary>
        public ExcelChart AddChartFromRange(string dataRangeA1, int row, int column, int widthPixels = 640, int heightPixels = 360,
            ExcelChartType type = ExcelChartType.ColumnClustered, bool hasHeaders = true, string? title = null, bool includeCachedData = true) {
            if (string.IsNullOrWhiteSpace(dataRangeA1)) throw new ArgumentNullException(nameof(dataRangeA1));
            if (!A1.TryParseRange(dataRangeA1, out int r1, out int c1, out int r2, out int c2)) {
                throw new ArgumentException($"Invalid A1 range '{dataRangeA1}'.", nameof(dataRangeA1));
            }

            int categoryCount = hasHeaders ? r2 - r1 : r2 - r1 + 1;
            int seriesCount = c2 - c1;
            if (categoryCount <= 0 || seriesCount <= 0) {
                throw new InvalidOperationException("Range must include at least one category row and one series column.");
            }

            var range = new ExcelChartDataRange(Name, r1, c1, categoryCount, seriesCount, hasHeaderRow: hasHeaders);
            ExcelChartData? data = includeCachedData ? ExcelChartUtils.TryReadChartData(this, range) : null;
            return AddChartInternal(range, row, column, widthPixels, heightPixels, type, data, title);
        }

        /// <summary>
        /// Adds a chart from a table name on this sheet.
        /// </summary>
        public ExcelChart AddChartFromTable(string tableName, int row, int column, int widthPixels = 640, int heightPixels = 360,
            ExcelChartType type = ExcelChartType.ColumnClustered, string? title = null, bool includeCachedData = true) {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));

            var table = _worksheetPart.TableDefinitionParts
                .Select(p => p.Table)
                .FirstOrDefault(t => string.Equals(t?.Name?.Value ?? t?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase));
            if (table?.Reference?.Value == null) {
                throw new InvalidOperationException($"Table '{tableName}' was not found on sheet '{Name}'.");
            }

            return AddChartFromRange(table.Reference.Value, row, column, widthPixels, heightPixels, type, hasHeaders: true, title, includeCachedData);
        }

        /// <summary>
        /// Adds a scatter chart using explicit X/Y ranges.
        /// </summary>
        public ExcelChart AddScatterChartFromRanges(IEnumerable<ExcelChartSeriesRange> seriesRanges, int row, int column, int widthPixels = 640,
            int heightPixels = 360, string? title = null) {
            return AddChartFromSeriesRanges(seriesRanges, row, column, widthPixels, heightPixels, ExcelChartType.Scatter, title);
        }

        /// <summary>
        /// Adds a bubble chart using explicit X/Y/size ranges.
        /// </summary>
        public ExcelChart AddBubbleChartFromRanges(IEnumerable<ExcelChartSeriesRange> seriesRanges, int row, int column, int widthPixels = 640,
            int heightPixels = 360, string? title = null) {
            return AddChartFromSeriesRanges(seriesRanges, row, column, widthPixels, heightPixels, ExcelChartType.Bubble, title);
        }

        private ExcelChart AddChartFromSeriesRanges(IEnumerable<ExcelChartSeriesRange> seriesRanges, int row, int column, int widthPixels,
            int heightPixels, ExcelChartType type, string? title) {
            if (seriesRanges == null) throw new ArgumentNullException(nameof(seriesRanges));
            if (row <= 0 || column <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (widthPixels <= 0 || heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));

            var ranges = seriesRanges.ToList();
            if (ranges.Count == 0) {
                throw new ArgumentException("At least one series range is required.", nameof(seriesRanges));
            }
            if (type == ExcelChartType.Bubble && ranges.Any(r => string.IsNullOrWhiteSpace(r.BubbleSizeRangeA1))) {
                throw new ArgumentException("Bubble charts require bubble size ranges for each series.", nameof(seriesRanges));
            }

            Xdr.GraphicFrame? frame = null;
            DrawingsPart? drawingPart = null;

            WriteLock(() => {
                var drawing = _worksheetPart.Worksheet.GetFirstChild<Drawing>();
                if (drawing == null) {
                    drawingPart = _worksheetPart.AddNewPart<DrawingsPart>();
                    drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                    string relId = _worksheetPart.GetIdOfPart(drawingPart);
                    _worksheetPart.Worksheet.Append(new Drawing { Id = relId });
                } else {
                    drawingPart = (DrawingsPart)_worksheetPart.GetPartById(drawing.Id!);
                    drawingPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
                }

                ChartPart chartPart = drawingPart!.AddNewPart<ChartPart>();
                ExcelChartUtils.PopulateChartFromSeriesRanges(chartPart, type, Name, ranges, title);
                if (_excelDocument.DefaultChartStylePreset != null) {
                    ExcelChartUtils.ApplyChartStyle(chartPart, _excelDocument.DefaultChartStylePreset);
                }
                string chartRelId = drawingPart.GetIdOfPart(chartPart);

                long cx = PxToEmu(widthPixels);
                long cy = PxToEmu(heightPixels);

                UInt32Value nvId = NextDrawingId(drawingPart);
                frame = new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvId, Name = $"Chart {nvId}" },
                        new Xdr.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true })
                    ),
                    new Xdr.Transform(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = cx, Cy = cy }
                    ),
                    new A.Graphic(
                        new A.GraphicData(
                            new C.ChartReference { Id = chartRelId }
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                );

                var anchor = new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId((column - 1).ToString()),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId((row - 1).ToString()),
                        new Xdr.RowOffset("0")
                    ),
                    new Xdr.Extent { Cx = cx, Cy = cy },
                    frame,
                    new Xdr.ClientData()
                );

                drawingPart.WorksheetDrawing.Append(anchor);
                drawingPart.WorksheetDrawing.Save();
                _worksheetPart.Worksheet.Save();
            });

            return new ExcelChart(frame!, drawingPart!, this);
        }

        private ExcelChart AddChartInternal(ExcelChartDataRange range, int row, int column, int widthPixels, int heightPixels,
            ExcelChartType type, ExcelChartData? data, string? title) {
            Xdr.GraphicFrame? frame = null;
            DrawingsPart? drawingPart = null;

            WriteLock(() => {
                var drawing = _worksheetPart.Worksheet.GetFirstChild<Drawing>();
                if (drawing == null) {
                    drawingPart = _worksheetPart.AddNewPart<DrawingsPart>();
                    drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                    string relId = _worksheetPart.GetIdOfPart(drawingPart);
                    _worksheetPart.Worksheet.Append(new Drawing { Id = relId });
                } else {
                    drawingPart = (DrawingsPart)_worksheetPart.GetPartById(drawing.Id!);
                    drawingPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
                }

                ChartPart chartPart = drawingPart!.AddNewPart<ChartPart>();
                ExcelChartUtils.PopulateChart(chartPart, type, range, data, title);
                if (_excelDocument.DefaultChartStylePreset != null) {
                    ExcelChartUtils.ApplyChartStyle(chartPart, _excelDocument.DefaultChartStylePreset);
                }
                string chartRelId = drawingPart.GetIdOfPart(chartPart);

                long cx = PxToEmu(widthPixels);
                long cy = PxToEmu(heightPixels);

                UInt32Value nvId = NextDrawingId(drawingPart);
                frame = new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvId, Name = $"Chart {nvId}" },
                        new Xdr.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true })
                    ),
                    new Xdr.Transform(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = cx, Cy = cy }
                    ),
                    new A.Graphic(
                        new A.GraphicData(
                            new C.ChartReference { Id = chartRelId }
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                );

                var anchor = new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId((column - 1).ToString()),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId((row - 1).ToString()),
                        new Xdr.RowOffset("0")
                    ),
                    new Xdr.Extent { Cx = cx, Cy = cy },
                    frame,
                    new Xdr.ClientData()
                );

                drawingPart.WorksheetDrawing.Append(anchor);
                drawingPart.WorksheetDrawing.Save();
                _worksheetPart.Worksheet.Save();
            });

            return new ExcelChart(frame!, drawingPart!, this, range);
        }
    }
}
