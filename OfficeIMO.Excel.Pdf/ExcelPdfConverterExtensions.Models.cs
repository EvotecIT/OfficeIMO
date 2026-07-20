using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private sealed class WorksheetPdfExportPlan {
            public WorksheetPdfExportPlan(string sheetName, ExcelSheetPageSetup? pageSetup, ExcelSheet.HeaderFooterSnapshot? headerFooter, SheetExportData exportData, IReadOnlyList<WorksheetImageExportData> images, IReadOnlyList<WorksheetChartExportData> charts, bool hasTable, int exportedRows, IReadOnlyList<int> manualRowBreaks, IReadOnlyList<int> manualColumnBreaks, string bookmarkName, WorksheetGeometryData geometry, bool clipToExportRange) {
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
            public ExcelSheet.HeaderFooterSnapshot? HeaderFooter { get; }
            public SheetExportData ExportData { get; }
            public IReadOnlyList<WorksheetImageExportData> Images { get; }
            public IReadOnlyList<WorksheetChartExportData> Charts { get; }
            public bool HasTable { get; }
            public int ExportedRows { get; }
            public IReadOnlyList<int> ManualRowBreaks { get; }
            public IReadOnlyList<int> ManualColumnBreaks { get; }
            public string BookmarkName { get; }
            public WorksheetGeometryData Geometry { get; }
            public bool ClipToExportRange { get; }
        }

        private sealed class SheetExportData {
            public SheetExportData(object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int headerRowCount, int firstBodyRowNumber, ConditionalFillData? conditionalFills = null) {
                Values = values;
                Styles = styles;
                Hyperlinks = hyperlinks;
                CellReferences = cellReferences;
                MergedCells = mergedCells;
                ColumnWidths = columnWidths;
                RowHeights = rowHeights;
                HeaderRowCount = headerRowCount;
                FirstBodyRowNumber = firstBodyRowNumber;
                ConditionalFills = conditionalFills;
            }

            public object?[,] Values { get; }
            public ExcelCellStyleSnapshot?[,]? Styles { get; }
            public ExcelHyperlinkSnapshot?[,]? Hyperlinks { get; }
            public string?[,]? CellReferences { get; }
            public MergeLayoutData? MergedCells { get; }
            public ColumnLayoutData? ColumnWidths { get; }
            public RowLayoutData? RowHeights { get; }
            public int HeaderRowCount { get; }
            public int FirstBodyRowNumber { get; }
            public ConditionalFillData? ConditionalFills { get; }
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
            public ConditionalIconCell(PdfCore.PdfCellIconKind kind, PdfCore.PdfColor color) {
                Kind = kind;
                Color = color;
            }

            public PdfCore.PdfCellIconKind Kind { get; }
            public PdfCore.PdfColor Color { get; }
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
                double defaultColumnWidth,
                double defaultRowHeight,
                IReadOnlyList<ExcelColumnSnapshot> columns,
                IReadOnlyList<ExcelRowSnapshot> rows,
                bool respectHiddenRowsAndColumns) {
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
                DefaultColumnWidth = defaultColumnWidth;
                DefaultRowHeight = defaultRowHeight;
                Columns = columns;
                Rows = rows;
                RespectHiddenRowsAndColumns = respectHiddenRowsAndColumns;
            }

            public int FirstRow { get; }
            public int FirstColumn { get; }
            public int LastRow { get; }
            public int LastColumn { get; }
            public double DefaultColumnWidth { get; }
            public double DefaultRowHeight { get; }
            public IReadOnlyList<ExcelColumnSnapshot> Columns { get; }
            public IReadOnlyList<ExcelRowSnapshot> Rows { get; }
            public bool RespectHiddenRowsAndColumns { get; }
        }

        private sealed class WorksheetChartExportData {
            public WorksheetChartExportData(ExcelChartSnapshot snapshot) {
                Snapshot = snapshot;
            }

            public ExcelChartSnapshot Snapshot { get; }
        }

        private sealed class RangeExportData {
            public RangeExportData(object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights) {
                Values = values;
                Styles = styles;
                Hyperlinks = hyperlinks;
                CellReferences = cellReferences;
                MergedCells = mergedCells;
                ColumnWidths = columnWidths;
                RowHeights = rowHeights;
            }

            public object?[,] Values { get; }
            public ExcelCellStyleSnapshot?[,]? Styles { get; }
            public ExcelHyperlinkSnapshot?[,]? Hyperlinks { get; }
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
            public TableChunk(IReadOnlyList<int> rowIndexes, int headerRowCount, int startColumn, int columnCount) {
                RowIndexes = rowIndexes;
                HeaderRowCount = headerRowCount;
                StartColumn = startColumn;
                ColumnCount = columnCount;
            }

            public IReadOnlyList<int> RowIndexes { get; }
            public int HeaderRowCount { get; }
            public int StartColumn { get; }
            public int ColumnCount { get; }
        }

        private sealed class MergeLayoutData {
            private readonly MergeSpan?[,] _spans;
            private readonly bool[,] _continuations;

            public MergeLayoutData(int rowCount, int columnCount) {
                _spans = new MergeSpan?[rowCount, columnCount];
                _continuations = new bool[rowCount, columnCount];
            }

            public bool HasAny { get; private set; }

            public void SetSpan(int row, int column, int rowSpan, int columnSpan) {
                if (row < 0 || column < 0 || row >= _spans.GetLength(0) || column >= _spans.GetLength(1)) {
                    return;
                }

                rowSpan = Math.Min(rowSpan, _spans.GetLength(0) - row);
                columnSpan = Math.Min(columnSpan, _spans.GetLength(1) - column);
                if (rowSpan <= 1 && columnSpan <= 1) {
                    return;
                }

                _spans[row, column] = new MergeSpan(rowSpan, columnSpan);
                for (int r = row; r < row + rowSpan; r++) {
                    for (int c = column; c < column + columnSpan; c++) {
                        if (r != row || c != column) {
                            _continuations[r, c] = true;
                        }
                    }
                }

                HasAny = true;
            }

            public MergeSpan? GetSpan(int row, int column) =>
                row >= 0 && column >= 0 && row < _spans.GetLength(0) && column < _spans.GetLength(1)
                    ? _spans[row, column]
                    : null;

            public bool IsContinuation(int row, int column) =>
                row >= 0 && column >= 0 && row < _continuations.GetLength(0) && column < _continuations.GetLength(1) && _continuations[row, column];

            public void CopyTo(MergeLayoutData target, int rowOffset) {
                for (int row = 0; row < _spans.GetLength(0); row++) {
                    for (int column = 0; column < _spans.GetLength(1); column++) {
                        MergeSpan? span = _spans[row, column];
                        if (span != null) {
                            target.SetSpan(row + rowOffset, column, span.RowSpan, span.ColumnSpan);
                        }
                    }
                }
            }
        }

        private sealed class MergeSpan {
            public MergeSpan(int rowSpan, int columnSpan) {
                RowSpan = rowSpan;
                ColumnSpan = columnSpan;
            }

            public int RowSpan { get; }
            public int ColumnSpan { get; }
        }
    }
}
