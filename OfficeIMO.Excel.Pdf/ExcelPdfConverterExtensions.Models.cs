using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private sealed class WorksheetPdfExportPlan {
            public WorksheetPdfExportPlan(string sheetName, ExcelSheetPageSetup? pageSetup, ExcelSheet.ExcelHeaderFooterInfo? headerFooter, SheetExportData exportData, IReadOnlyList<WorksheetImageExportData> images, IReadOnlyList<WorksheetChartExportData> charts, bool hasTable, int exportedRows, IReadOnlyList<int> manualRowBreaks, IReadOnlyList<int> manualColumnBreaks, string bookmarkName, WorksheetGeometryData geometry, bool clipToExportRange) {
                SheetName = sheetName;
                PageSetup = pageSetup;
                HeaderFooter = headerFooter;
                ExportData = exportData;
                Images = images;
                Charts = charts;
                HasTable = hasTable;
                ExportedRows = exportedRows;
                ManualRowBreaks = manualRowBreaks;
                ManualColumnBreaks = manualColumnBreaks;
                BookmarkName = bookmarkName;
                Geometry = geometry;
                ClipToExportRange = clipToExportRange;
            }

            public string SheetName { get; }
            public ExcelSheetPageSetup? PageSetup { get; }
            public ExcelSheet.ExcelHeaderFooterInfo? HeaderFooter { get; }
            public SheetExportData ExportData { get; }
            public IReadOnlyList<WorksheetImageExportData> Images { get; }
            public IReadOnlyList<WorksheetChartExportData> Charts { get; }
            public ExcelSheetPrintOptions? PrintOptions { get; set; }
            public IReadOnlyList<OfficeIMO.Excel.Utilities.ExcelWorksheetDrawingObjectInfo> DrawingObjects { get; set; } = Array.Empty<OfficeIMO.Excel.Utilities.ExcelWorksheetDrawingObjectInfo>();
            public bool HasTable { get; }
            public int ExportedRows { get; }
            public IReadOnlyList<int> ManualRowBreaks { get; }
            public IReadOnlyList<int> ManualColumnBreaks { get; }
            public string BookmarkName { get; }
            public WorksheetGeometryData Geometry { get; }
            public bool ClipToExportRange { get; }
            public WorksheetTextMeasurer? TextMeasurer { get; set; }
            public Dictionary<WorksheetChartExportData, OfficeChartRenderingResult> ChartRenderings { get; } = new();
        }

        private sealed class SheetExportData {
            public SheetExportData(object?[,] values, ExcelCellStyleInfo?[,]? styles, ExcelHyperlinkInfo?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int headerRowCount, int firstBodyRowNumber, IReadOnlyList<StructuredTableVisualData> structuredTables, ConditionalFillData? conditionalFills = null, IReadOnlyList<int>? printTitleColumnIndexes = null, IReadOnlyList<int>? printTitleRowIndexes = null) {
                Values = values;
                Styles = styles;
                Hyperlinks = hyperlinks;
                CellReferences = cellReferences;
                MergedCells = mergedCells;
                ColumnWidths = columnWidths;
                RowHeights = rowHeights;
                HeaderRowCount = headerRowCount;
                FirstBodyRowNumber = firstBodyRowNumber;
                StructuredTables = structuredTables;
                ConditionalFills = conditionalFills;
                PrintTitleColumnIndexes = printTitleColumnIndexes ?? Array.Empty<int>();
                RepeatingRowIndexes = Enumerable.Range(0, Math.Min(Math.Max(0, headerRowCount), values.GetLength(0)))
                    .Concat(printTitleRowIndexes ?? Array.Empty<int>()).Distinct().OrderBy(row => row).ToArray();
            }

            public object?[,] Values { get; }
            public ExcelCellStyleInfo?[,]? Styles { get; }
            public ExcelHyperlinkInfo?[,]? Hyperlinks { get; }
            public string?[,]? CellReferences { get; }
            public MergeLayoutData? MergedCells { get; }
            public ColumnLayoutData? ColumnWidths { get; }
            public RowLayoutData? RowHeights { get; }
            public int HeaderRowCount { get; }
            public int FirstBodyRowNumber { get; }
            public IReadOnlyList<StructuredTableVisualData> StructuredTables { get; }
            public ConditionalFillData? ConditionalFills { get; }
            public IReadOnlyList<int> PrintTitleColumnIndexes { get; }
            public IReadOnlyList<int> RepeatingRowIndexes { get; }
        }

        private sealed class ConditionalFillData {
            public ConditionalFillData(IReadOnlyDictionary<(int Row, int Column), string> fillColors, IReadOnlyDictionary<(int Row, int Column), ConditionalDataBarCell> dataBars, IReadOnlyDictionary<(int Row, int Column), ConditionalIconCell> icons) {
                FillColors = fillColors;
                DataBars = dataBars;
                Icons = icons;
            }

            public IReadOnlyDictionary<(int Row, int Column), string> FillColors { get; }
            public IReadOnlyDictionary<(int Row, int Column), ConditionalDataBarCell> DataBars { get; }
            public IReadOnlyDictionary<(int Row, int Column), ConditionalIconCell> Icons { get; }
        }

        private sealed class ConditionalDataBarCell {
            public ConditionalDataBarCell(string color, double startRatio, double ratio) {
                Color = color;
                StartRatio = startRatio;
                Ratio = ratio;
            }

            public string Color { get; }
            public double StartRatio { get; }
            public double Ratio { get; }
        }

        private sealed class ConditionalIconCell {
            public ConditionalIconCell(OfficeConditionalIconKind kind, bool showValue) {
                Kind = kind;
                ShowValue = showValue;
            }

            public OfficeConditionalIconKind Kind { get; }
            public bool ShowValue { get; }
        }

        private sealed class WorksheetImageExportData {
            public WorksheetImageExportData(byte[] bytes, double widthPoints, double heightPoints, string cellReference, int rowIndex, int columnIndex, double offsetXPoints, double offsetYPoints, double rotationDegrees, bool horizontalFlip, bool verticalFlip, string? alternativeText) {
                Bytes = bytes;
                WidthPoints = widthPoints;
                HeightPoints = heightPoints;
                CellReference = cellReference;
                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
                OffsetXPoints = offsetXPoints;
                OffsetYPoints = offsetYPoints;
                RotationDegrees = rotationDegrees;
                HorizontalFlip = horizontalFlip;
                VerticalFlip = verticalFlip;
                AlternativeText = alternativeText;
            }

            public byte[] Bytes { get; }
            public double WidthPoints { get; }
            public double HeightPoints { get; }
            public string CellReference { get; }
            public int RowIndex { get; }
            public int ColumnIndex { get; }
            public double OffsetXPoints { get; }
            public double OffsetYPoints { get; }
            public double RotationDegrees { get; }
            public bool HorizontalFlip { get; }
            public bool VerticalFlip { get; }
            public string? AlternativeText { get; }
        }

        private sealed class WorksheetGeometryData {
            public WorksheetGeometryData(
                int firstRow,
                int firstColumn,
                int lastRow,
                int lastColumn,
                double? defaultColumnWidth,
                double baseColumnWidth,
                double defaultRowHeight,
                IReadOnlyList<ExcelColumnInfo> columns,
                IReadOnlyList<ExcelRowInfo> rows,
                bool useWorksheetColumnWidths,
                bool useWorksheetRowHeights,
                bool respectHiddenRowsAndColumns) {
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
                DefaultColumnWidth = defaultColumnWidth ?? 8.43D;
                HasExplicitDefaultColumnWidth = defaultColumnWidth.HasValue;
                BaseColumnWidth = baseColumnWidth;
                DefaultRowHeight = defaultRowHeight;
                Columns = columns;
                Rows = rows;
                UseWorksheetColumnWidths = useWorksheetColumnWidths;
                UseWorksheetRowHeights = useWorksheetRowHeights;
                RespectHiddenRowsAndColumns = respectHiddenRowsAndColumns;
            }

            public int FirstRow { get; }
            public int FirstColumn { get; }
            public int LastRow { get; }
            public int LastColumn { get; }
            public double DefaultColumnWidth { get; }
            public bool HasExplicitDefaultColumnWidth { get; }
            public double BaseColumnWidth { get; }
            public double? MaximumDigitWidthPoints { get; set; }
            /// <summary>Excel's automatic heights for rows without a stored height, keyed by source row.</summary>
            public Dictionary<int, double>? AutoRowHeights { get; set; }
            public double DefaultRowHeight { get; }
            public IReadOnlyList<ExcelColumnInfo> Columns { get; }
            public IReadOnlyList<ExcelRowInfo> Rows { get; }
            public bool UseWorksheetColumnWidths { get; }
            public bool UseWorksheetRowHeights { get; }
            public bool RespectHiddenRowsAndColumns { get; }
        }

        private sealed class WorksheetChartExportData {
            public WorksheetChartExportData(ExcelChartInfo snapshot) {
                Snapshot = snapshot;
            }

            public ExcelChartInfo Snapshot { get; }
        }

        private sealed class RangeExportData {
            public RangeExportData(object?[,] values, ExcelCellStyleInfo?[,]? styles, ExcelHyperlinkInfo?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights) {
                Values = values;
                Styles = styles;
                Hyperlinks = hyperlinks;
                CellReferences = cellReferences;
                MergedCells = mergedCells;
                ColumnWidths = columnWidths;
                RowHeights = rowHeights;
            }

            public object?[,] Values { get; }
            public ExcelCellStyleInfo?[,]? Styles { get; }
            public ExcelHyperlinkInfo?[,]? Hyperlinks { get; }
            public string?[,]? CellReferences { get; }
            public MergeLayoutData? MergedCells { get; }
            public ColumnLayoutData? ColumnWidths { get; }
            public RowLayoutData? RowHeights { get; }
        }

        private sealed class VisibilityLayoutData {
            public VisibilityLayoutData(List<int> rowOffsets, List<int> columnOffsets, int originalRowCount, int originalColumnCount) {
                RowOffsets = rowOffsets;
                ColumnOffsets = columnOffsets;
                OriginalRowCount = originalRowCount;
                OriginalColumnCount = originalColumnCount;
            }

            public List<int> RowOffsets { get; }
            public List<int> ColumnOffsets { get; }
            public int OriginalRowCount { get; }
            public int OriginalColumnCount { get; }
        }

        private sealed class ColumnLayoutData {
            public ColumnLayoutData(List<double> widthWeights, double approximateWidthPoints) {
                WidthWeights = widthWeights;
                ApproximateWidthPoints = approximateWidthPoints;
            }

            public List<double> WidthWeights { get; }
            public double ApproximateWidthPoints { get; }
        }

        private sealed class RowLayoutData {
            public RowLayoutData(List<double?> minHeights) {
                MinHeights = minHeights;
            }

            public List<double?> MinHeights { get; }
        }

        private sealed class TableAxisChunk {
            public TableAxisChunk(int start, int count) {
                Start = start;
                Count = count;
            }

            public int Start { get; }
            public int Count { get; }
        }

        private sealed class TableChunk {
            public TableChunk(IReadOnlyList<int> rowIndexes, int headerRowCount, int startColumn, int columnCount, IReadOnlyList<int>? columnIndexes = null) {
                RowIndexes = rowIndexes;
                HeaderRowCount = headerRowCount;
                StartColumn = startColumn;
                ColumnCount = columnCount;
                ColumnIndexes = columnIndexes ?? Enumerable.Range(startColumn, columnCount).ToArray();
            }

            public IReadOnlyList<int> RowIndexes { get; }
            public int HeaderRowCount { get; }
            public int StartColumn { get; }
            public int ColumnCount { get; }
            public IReadOnlyList<int> ColumnIndexes { get; }
        }

        private sealed class MergeLayoutData {
            private readonly int[,] _mergeIndexes;
            private readonly List<MergeRegion> _regions = new();

            public MergeLayoutData(int rowCount, int columnCount) {
                _mergeIndexes = new int[rowCount, columnCount];
            }

            public bool HasAny { get; private set; }

            public void SetSpan(int row, int column, int rowSpan, int columnSpan, PdfCore.PdfTableCellViewport? textViewport = null, string? imageAnchorReference = null) {
                if (row < 0 || column < 0 || row >= _mergeIndexes.GetLength(0) || column >= _mergeIndexes.GetLength(1)) {
                    return;
                }

                rowSpan = Math.Min(rowSpan, _mergeIndexes.GetLength(0) - row);
                columnSpan = Math.Min(columnSpan, _mergeIndexes.GetLength(1) - column);
                if (rowSpan <= 0 || columnSpan <= 0 || rowSpan == 1 && columnSpan == 1 && textViewport == null) {
                    return;
                }

                _regions.Add(new MergeRegion(row, column, new MergeSpan(rowSpan, columnSpan, textViewport, imageAnchorReference)));
                for (int r = row; r < row + rowSpan; r++) {
                    for (int c = column; c < column + columnSpan; c++) {
                        _mergeIndexes[r, c] = _regions.Count;
                    }
                }

                HasAny = true;
            }

            public MergeRegion? GetRegion(int row, int column) {
                if (row < 0 || column < 0 || row >= _mergeIndexes.GetLength(0) || column >= _mergeIndexes.GetLength(1)) return null;
                int index = _mergeIndexes[row, column];
                return index == 0 ? null : _regions[index - 1];
            }

            public MergeSpan? GetSpan(int row, int column) => GetRegion(row, column) is { } region && region.Row == row && region.Column == column ? region.Span : null;
            public bool IsContinuation(int row, int column) => GetRegion(row, column) is { } region && (region.Row != row || region.Column != column);

            public void CopyTo(MergeLayoutData target, int rowOffset, int columnOffset = 0) {
                foreach (MergeRegion region in _regions)
                    target.SetSpan(region.Row + rowOffset, region.Column + columnOffset, region.Span.RowSpan, region.Span.ColumnSpan, region.Span.Viewport, region.Span.ImageAnchorReference);
            }
        }

        private sealed class MergeRegion {
            public MergeRegion(int row, int column, MergeSpan span) { Row = row; Column = column; Span = span; }
            public int Row { get; }
            public int Column { get; }
            public MergeSpan Span { get; }
        }

        private sealed class MergeSpan {
            public MergeSpan(int rowSpan, int columnSpan, PdfCore.PdfTableCellViewport? textViewport = null, string? imageAnchorReference = null) {
                RowSpan = rowSpan;
                ColumnSpan = columnSpan;
                Viewport = textViewport;
                ImageAnchorReference = imageAnchorReference;
            }

            public int RowSpan { get; }
            public int ColumnSpan { get; }
            public PdfCore.PdfTableCellViewport? Viewport { get; }
            public string? ImageAnchorReference { get; }
        }
    }
}
