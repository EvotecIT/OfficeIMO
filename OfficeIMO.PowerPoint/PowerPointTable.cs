using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public class PowerPointTable : PowerPointShape {
        private const int EmusPerPoint = 12700;
        internal PowerPointTable(GraphicFrame frame) : base(frame) {
        }

        private GraphicFrame Frame => (GraphicFrame)Element;
        internal A.Table TableElement => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

        /// <summary>
        ///     Returns number of rows in the table.
        /// </summary>
        public int Rows => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.Elements<A.TableRow>().Count();

        /// <summary>
        ///     Returns number of columns in the table.
        /// </summary>
        public int Columns => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.TableGrid!.Elements<A.GridColumn>()
            .Count();

        /// <summary>
        ///     Row wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableRow> RowItems =>
            TableElement.Elements<A.TableRow>().Select(r => new PowerPointTableRow(this, r)).ToList();

        /// <summary>
        ///     Column wrappers for the table.
        /// </summary>
        public IReadOnlyList<PowerPointTableColumn> ColumnItems =>
            TableElement.TableGrid?.Elements<A.GridColumn>()
                .Select(c => new PowerPointTableColumn(this, c))
                .ToList() ?? new List<PowerPointTableColumn>();

        /// <summary>
        ///     Enables or disables header row styling (firstRow attribute) on the table.
        /// </summary>
        public bool HeaderRow {
            get => FirstRow;
            set => FirstRow = value;
        }

        /// <summary>
        ///     Enables or disables first row styling on the table.
        /// </summary>
        public bool FirstRow {
            get => TableElement.TableProperties?.FirstRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.FirstRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables last row styling on the table.
        /// </summary>
        public bool LastRow {
            get => TableElement.TableProperties?.LastRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.LastRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables first column styling on the table.
        /// </summary>
        public bool FirstColumn {
            get => TableElement.TableProperties?.FirstColumn?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.FirstColumn = value;
            }
        }

        /// <summary>
        ///     Enables or disables last column styling on the table.
        /// </summary>
        public bool LastColumn {
            get => TableElement.TableProperties?.LastColumn?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.LastColumn = value;
            }
        }

        /// <summary>
        ///     Enables or disables banded rows styling (bandRow attribute) on the table.
        /// </summary>
        public bool BandedRows {
            get => TableElement.TableProperties?.BandRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.BandRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables banded columns styling (bandCol attribute) on the table.
        /// </summary>
        public bool BandedColumns {
            get => TableElement.TableProperties?.BandColumn?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.BandColumn = value;
            }
        }

        /// <summary>
        ///     Gets or sets the table style ID GUID.
        /// </summary>
        public string? StyleId {
            get => TableElement.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.RemoveAllChildren<A.TableStyleId>();
                if (!string.IsNullOrWhiteSpace(value)) {
                    TableElement.TableProperties.Append(new A.TableStyleId { Text = value! });
                }
            }
        }

        /// <summary>
        ///     Sets the width of a specific column in EMUs.
        /// </summary>
        public void SetColumnWidth(int columnIndex, long width) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            column.Width = width;
        }

        /// <summary>
        ///     Gets the width of a specific column in EMUs.
        /// </summary>
        public long GetColumnWidth(int columnIndex) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            return column.Width?.Value ?? 0L;
        }

        /// <summary>
        ///     Gets the width of a specific column in points.
        /// </summary>
        public double GetColumnWidthPoints(int columnIndex) {
            return PowerPointUnits.ToPoints(GetColumnWidth(columnIndex));
        }

        /// <summary>
        ///     Gets the width of a specific column in centimeters.
        /// </summary>
        public double GetColumnWidthCm(int columnIndex) {
            return PowerPointUnits.ToCentimeters(GetColumnWidth(columnIndex));
        }

        /// <summary>
        ///     Gets the width of a specific column in inches.
        /// </summary>
        public double GetColumnWidthInches(int columnIndex) {
            return PowerPointUnits.ToInches(GetColumnWidth(columnIndex));
        }

        /// <summary>
        ///     Sets the width of a specific column in points.
        /// </summary>
        public void SetColumnWidthPoints(int columnIndex, double widthPoints) {
            SetColumnWidth(columnIndex, ToEmus(widthPoints));
        }

        /// <summary>
        ///     Sets the width of a specific column in centimeters.
        /// </summary>
        public void SetColumnWidthCm(int columnIndex, double widthCm) {
            SetColumnWidth(columnIndex, PowerPointUnits.FromCentimeters(widthCm));
        }

        /// <summary>
        ///     Sets the width of a specific column in inches.
        /// </summary>
        public void SetColumnWidthInches(int columnIndex, double widthInches) {
            SetColumnWidth(columnIndex, PowerPointUnits.FromInches(widthInches));
        }

        /// <summary>
        ///     Sets widths for columns in points (applies to the first N columns provided).
        /// </summary>
        public void SetColumnWidthsPoints(params double[] widthsPoints) {
            if (widthsPoints == null) {
                throw new ArgumentNullException(nameof(widthsPoints));
            }
            int count = Math.Min(widthsPoints.Length, Columns);
            for (int i = 0; i < count; i++) {
                SetColumnWidthPoints(i, widthsPoints[i]);
            }
        }

        /// <summary>
        ///     Sets widths for columns in centimeters (applies to the first N columns provided).
        /// </summary>
        public void SetColumnWidthsCm(params double[] widthsCm) {
            if (widthsCm == null) {
                throw new ArgumentNullException(nameof(widthsCm));
            }
            int count = Math.Min(widthsCm.Length, Columns);
            for (int i = 0; i < count; i++) {
                SetColumnWidthCm(i, widthsCm[i]);
            }
        }

        /// <summary>
        ///     Sets widths for columns in inches (applies to the first N columns provided).
        /// </summary>
        public void SetColumnWidthsInches(params double[] widthsInches) {
            if (widthsInches == null) {
                throw new ArgumentNullException(nameof(widthsInches));
            }
            int count = Math.Min(widthsInches.Length, Columns);
            for (int i = 0; i < count; i++) {
                SetColumnWidthInches(i, widthsInches[i]);
            }
        }

        /// <summary>
        ///     Sets column widths proportionally using ratios.
        /// </summary>
        public void SetColumnWidthsByRatio(params double[] ratios) {
            if (ratios == null) {
                throw new ArgumentNullException(nameof(ratios));
            }
            if (ratios.Length == 0) {
                throw new ArgumentException("At least one ratio is required.", nameof(ratios));
            }

            int count = Math.Min(ratios.Length, Columns);
            double totalRatio = 0;
            for (int i = 0; i < count; i++) {
                if (ratios[i] <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(ratios), "Ratios must be positive.");
                }
                totalRatio += ratios[i];
            }

            long totalWidth = Width;
            if (totalWidth <= 0) {
                totalWidth = TableElement.TableGrid?.Elements<A.GridColumn>()
                    .Sum(c => c.Width?.Value ?? 0) ?? 0;
            }
            if (totalWidth <= 0) {
                throw new InvalidOperationException("Table width is not available.");
            }

            long assigned = 0;
            for (int i = 0; i < count; i++) {
                long width = (long)Math.Round(totalWidth * (ratios[i] / totalRatio));
                SetColumnWidth(i, width);
                assigned += width;
            }

            if (assigned != totalWidth && count > 0) {
                int adjustIndex = count - 1;
                A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(adjustIndex);
                column.Width = (column.Width ?? 0) + (totalWidth - assigned);
            }
        }

        /// <summary>
        ///     Sets column widths evenly across the table width.
        /// </summary>
        public void SetColumnWidthsEvenly() {
            if (Columns == 0) {
                return;
            }
            SetColumnWidthsByRatio(Enumerable.Repeat(1d, Columns).ToArray());
        }

        /// <summary>
        ///     Sets the height of a specific row in EMUs.
        /// </summary>
        public void SetRowHeight(int rowIndex, long height) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            row.Height = height;
        }

        /// <summary>
        ///     Gets the height of a specific row in EMUs.
        /// </summary>
        public long GetRowHeight(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            return row.Height?.Value ?? 0L;
        }

        /// <summary>
        ///     Gets the height of a specific row in points.
        /// </summary>
        public double GetRowHeightPoints(int rowIndex) {
            return PowerPointUnits.ToPoints(GetRowHeight(rowIndex));
        }

        /// <summary>
        ///     Gets the height of a specific row in centimeters.
        /// </summary>
        public double GetRowHeightCm(int rowIndex) {
            return PowerPointUnits.ToCentimeters(GetRowHeight(rowIndex));
        }

        /// <summary>
        ///     Gets the height of a specific row in inches.
        /// </summary>
        public double GetRowHeightInches(int rowIndex) {
            return PowerPointUnits.ToInches(GetRowHeight(rowIndex));
        }

        /// <summary>
        ///     Sets the height of a specific row in points.
        /// </summary>
        public void SetRowHeightPoints(int rowIndex, double heightPoints) {
            SetRowHeight(rowIndex, ToEmus(heightPoints));
        }

        /// <summary>
        ///     Sets the height of a specific row in centimeters.
        /// </summary>
        public void SetRowHeightCm(int rowIndex, double heightCm) {
            SetRowHeight(rowIndex, PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Sets the height of a specific row in inches.
        /// </summary>
        public void SetRowHeightInches(int rowIndex, double heightInches) {
            SetRowHeight(rowIndex, PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Sets heights for rows in points (applies to the first N rows provided).
        /// </summary>
        public void SetRowHeightsPoints(params double[] heightsPoints) {
            if (heightsPoints == null) {
                throw new ArgumentNullException(nameof(heightsPoints));
            }
            int count = Math.Min(heightsPoints.Length, Rows);
            for (int i = 0; i < count; i++) {
                SetRowHeightPoints(i, heightsPoints[i]);
            }
        }

        /// <summary>
        ///     Sets heights for rows in centimeters (applies to the first N rows provided).
        /// </summary>
        public void SetRowHeightsCm(params double[] heightsCm) {
            if (heightsCm == null) {
                throw new ArgumentNullException(nameof(heightsCm));
            }
            int count = Math.Min(heightsCm.Length, Rows);
            for (int i = 0; i < count; i++) {
                SetRowHeightCm(i, heightsCm[i]);
            }
        }

        /// <summary>
        ///     Sets heights for rows in inches (applies to the first N rows provided).
        /// </summary>
        public void SetRowHeightsInches(params double[] heightsInches) {
            if (heightsInches == null) {
                throw new ArgumentNullException(nameof(heightsInches));
            }
            int count = Math.Min(heightsInches.Length, Rows);
            for (int i = 0; i < count; i++) {
                SetRowHeightInches(i, heightsInches[i]);
            }
        }

        /// <summary>
        ///     Sets row heights evenly across the table height.
        /// </summary>
        public void SetRowHeightsEvenly() {
            if (Rows == 0) {
                return;
            }

            long totalHeight = Height;
            if (totalHeight <= 0) {
                totalHeight = TableElement.Elements<A.TableRow>()
                    .Sum(r => r.Height?.Value ?? 0L);
            }
            if (totalHeight <= 0) {
                return;
            }

            long rowHeight = (long)Math.Floor(totalHeight / (double)Rows);
            long assigned = 0;
            for (int i = 0; i < Rows; i++) {
                SetRowHeight(i, rowHeight);
                assigned += rowHeight;
            }

            if (assigned != totalHeight && Rows > 0) {
                int adjustIndex = Rows - 1;
                A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(adjustIndex);
                row.Height = (row.Height ?? 0) + (totalHeight - assigned);
            }
        }

        /// <summary>
        ///     Sets row heights proportionally using ratios.
        /// </summary>
        public void SetRowHeightsByRatio(params double[] ratios) {
            if (ratios == null) {
                throw new ArgumentNullException(nameof(ratios));
            }
            if (ratios.Length == 0) {
                throw new ArgumentException("At least one ratio is required.", nameof(ratios));
            }

            int count = Math.Min(ratios.Length, Rows);
            double totalRatio = 0;
            for (int i = 0; i < count; i++) {
                if (ratios[i] <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(ratios), "Ratios must be positive.");
                }
                totalRatio += ratios[i];
            }

            long totalHeight = Height;
            if (totalHeight <= 0) {
                totalHeight = TableElement.Elements<A.TableRow>()
                    .Sum(r => r.Height?.Value ?? 0L);
            }
            if (totalHeight <= 0) {
                throw new InvalidOperationException("Table height is not available.");
            }

            long assigned = 0;
            for (int i = 0; i < count; i++) {
                long height = (long)Math.Round(totalHeight * (ratios[i] / totalRatio));
                SetRowHeight(i, height);
                assigned += height;
            }

            if (assigned != totalHeight && count > 0) {
                int adjustIndex = count - 1;
                A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(adjustIndex);
                row.Height = (row.Height ?? 0) + (totalHeight - assigned);
            }
        }

        /// <summary>
        ///     Applies a style preset to the table.
        /// </summary>
        public void ApplyStyle(PowerPointTableStylePreset preset) {
            if (!string.IsNullOrWhiteSpace(preset.StyleId)) {
                StyleId = preset.StyleId;
            }
            if (preset.FirstRow.HasValue) {
                FirstRow = preset.FirstRow.Value;
            }
            if (preset.LastRow.HasValue) {
                LastRow = preset.LastRow.Value;
            }
            if (preset.FirstColumn.HasValue) {
                FirstColumn = preset.FirstColumn.Value;
            }
            if (preset.LastColumn.HasValue) {
                LastColumn = preset.LastColumn.Value;
            }
            if (preset.BandedRows.HasValue) {
                BandedRows = preset.BandedRows.Value;
            }
            if (preset.BandedColumns.HasValue) {
                BandedColumns = preset.BandedColumns.Value;
            }
        }

        /// <summary>
        ///     Retrieves a cell at the specified row and column index.
        /// </summary>
        /// <param name="row">Zero-based row index.</param>
        /// <param name="column">Zero-based column index.</param>
        public PowerPointTableCell GetCell(int row, int column) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableRow tableRow = table.Elements<A.TableRow>().ElementAt(row);
            A.TableCell cell = tableRow.Elements<A.TableCell>().ElementAt(column);
            return new PowerPointTableCell(cell);
        }

        /// <summary>
        ///     Applies an action to all cells in the table.
        /// </summary>
        public void ApplyToCells(Action<PowerPointTableCell> action) {
            if (action == null) {
                throw new ArgumentNullException(nameof(action));
            }
            if (Rows == 0 || Columns == 0) {
                return;
            }

            ApplyToCells(0, Rows - 1, 0, Columns - 1, action);
        }

        /// <summary>
        ///     Applies an action to a rectangular range of cells.
        /// </summary>
        public void ApplyToCells(int startRow, int endRow, int startColumn, int endColumn, Action<PowerPointTableCell> action) {
            if (action == null) {
                throw new ArgumentNullException(nameof(action));
            }
            if (Rows == 0 || Columns == 0) {
                return;
            }

            int topRow = Math.Min(startRow, endRow);
            int bottomRow = Math.Max(startRow, endRow);
            int leftColumn = Math.Min(startColumn, endColumn);
            int rightColumn = Math.Max(startColumn, endColumn);

            if (topRow < 0 || leftColumn < 0) {
                throw new ArgumentOutOfRangeException("Row and column indices must be non-negative.");
            }
            if (bottomRow >= Rows || rightColumn >= Columns) {
                throw new ArgumentOutOfRangeException("Range exceeds table bounds.");
            }

            for (int r = topRow; r <= bottomRow; r++) {
                for (int c = leftColumn; c <= rightColumn; c++) {
                    action(GetCell(r, c));
                }
            }
        }

        /// <summary>
        ///     Applies an action to a specific row.
        /// </summary>
        public void ApplyToRow(int rowIndex, Action<PowerPointTableCell> action) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            ApplyToCells(rowIndex, rowIndex, 0, Columns - 1, action);
        }

        /// <summary>
        ///     Applies an action to a specific column.
        /// </summary>
        public void ApplyToColumn(int columnIndex, Action<PowerPointTableCell> action) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            ApplyToCells(0, Rows - 1, columnIndex, columnIndex, action);
        }

        /// <summary>
        ///     Sets cell padding in points for all cells.
        /// </summary>
        public void SetCellPaddingPoints(double? left, double? top, double? right, double? bottom) {
            ApplyToCells(cell => {
                cell.PaddingLeftPoints = left;
                cell.PaddingTopPoints = top;
                cell.PaddingRightPoints = right;
                cell.PaddingBottomPoints = bottom;
            });
        }

        /// <summary>
        ///     Sets cell padding in points for a range of cells.
        /// </summary>
        public void SetCellPaddingPoints(double? left, double? top, double? right, double? bottom,
            int startRow, int endRow, int startColumn, int endColumn) {
            ApplyToCells(startRow, endRow, startColumn, endColumn, cell => {
                cell.PaddingLeftPoints = left;
                cell.PaddingTopPoints = top;
                cell.PaddingRightPoints = right;
                cell.PaddingBottomPoints = bottom;
            });
        }

        /// <summary>
        ///     Sets cell padding in centimeters for all cells.
        /// </summary>
        public void SetCellPaddingCm(double? leftCm, double? topCm, double? rightCm, double? bottomCm) {
            ApplyToCells(cell => {
                cell.PaddingLeftCm = leftCm;
                cell.PaddingTopCm = topCm;
                cell.PaddingRightCm = rightCm;
                cell.PaddingBottomCm = bottomCm;
            });
        }

        /// <summary>
        ///     Sets cell padding in inches for all cells.
        /// </summary>
        public void SetCellPaddingInches(double? leftInches, double? topInches, double? rightInches, double? bottomInches) {
            ApplyToCells(cell => {
                cell.PaddingLeftInches = leftInches;
                cell.PaddingTopInches = topInches;
                cell.PaddingRightInches = rightInches;
                cell.PaddingBottomInches = bottomInches;
            });
        }

        /// <summary>
        ///     Sets cell alignment for all cells.
        /// </summary>
        public void SetCellAlignment(A.TextAlignmentTypeValues? horizontal, A.TextAnchoringTypeValues? vertical) {
            ApplyToCells(cell => {
                cell.HorizontalAlignment = horizontal;
                cell.VerticalAlignment = vertical;
            });
        }

        /// <summary>
        ///     Sets cell alignment for a range of cells.
        /// </summary>
        public void SetCellAlignment(A.TextAlignmentTypeValues? horizontal, A.TextAnchoringTypeValues? vertical,
            int startRow, int endRow, int startColumn, int endColumn) {
            ApplyToCells(startRow, endRow, startColumn, endColumn, cell => {
                cell.HorizontalAlignment = horizontal;
                cell.VerticalAlignment = vertical;
            });
        }

        /// <summary>
        ///     Applies borders to all cells.
        /// </summary>
        public void SetCellBorders(TableCellBorders borders, string color, double? widthPoints = null) {
            ApplyToCells(cell => cell.SetBorders(borders, color, widthPoints));
        }

        /// <summary>
        ///     Applies dashed borders to all cells.
        /// </summary>
        public void SetCellBorders(TableCellBorders borders, string color, double? widthPoints, A.PresetLineDashValues dash) {
            ApplyToCells(cell => cell.SetBorders(borders, color, widthPoints, dash));
        }

        /// <summary>
        ///     Clears borders for all cells.
        /// </summary>
        public void ClearCellBorders(TableCellBorders borders) {
            ApplyToCells(cell => cell.ClearBorders(borders));
        }

        /// <summary>
        ///     Retrieves a row wrapper at the specified index.
        /// </summary>
        public PowerPointTableRow GetRow(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }

            A.TableRow tableRow = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            return new PowerPointTableRow(this, tableRow);
        }

        /// <summary>
        ///     Retrieves a column wrapper at the specified index.
        /// </summary>
        public PowerPointTableColumn GetColumn(int columnIndex) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }

            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            return new PowerPointTableColumn(this, column);
        }

        /// <summary>
        ///     Merges a rectangular range of cells into the top-left cell.
        /// </summary>
        /// <param name="startRow">Zero-based start row.</param>
        /// <param name="startColumn">Zero-based start column.</param>
        /// <param name="endRow">Zero-based end row.</param>
        /// <param name="endColumn">Zero-based end column.</param>
        /// <param name="clearMergedContent">Whether to clear text from merged cells.</param>
        public void MergeCells(int startRow, int startColumn, int endRow, int endColumn, bool clearMergedContent = true) {
            int topRow = Math.Min(startRow, endRow);
            int bottomRow = Math.Max(startRow, endRow);
            int leftColumn = Math.Min(startColumn, endColumn);
            int rightColumn = Math.Max(startColumn, endColumn);

            if (topRow < 0 || leftColumn < 0) {
                throw new ArgumentOutOfRangeException("Row and column indices must be non-negative.");
            }
            if (bottomRow >= Rows || rightColumn >= Columns) {
                throw new ArgumentOutOfRangeException("Merge range exceeds table bounds.");
            }

            int rowSpan = bottomRow - topRow + 1;
            int colSpan = rightColumn - leftColumn + 1;
            if (rowSpan == 1 && colSpan == 1) {
                return;
            }

            A.Table table = TableElement;
            for (int r = topRow; r <= bottomRow; r++) {
                A.TableRow row = table.Elements<A.TableRow>().ElementAt(r);
                for (int c = leftColumn; c <= rightColumn; c++) {
                    A.TableCell cell = row.Elements<A.TableCell>().ElementAt(c);
                    bool isAnchor = r == topRow && c == leftColumn;

                    if (isAnchor) {
                        cell.RowSpan = rowSpan > 1 ? rowSpan : null;
                        cell.GridSpan = colSpan > 1 ? colSpan : null;
                        cell.HorizontalMerge = null;
                        cell.VerticalMerge = null;
                        continue;
                    }

                    cell.RowSpan = null;
                    cell.GridSpan = null;
                    cell.HorizontalMerge = c > leftColumn ? true : (bool?)null;
                    cell.VerticalMerge = r > topRow ? true : (bool?)null;

                    if (clearMergedContent) {
                        ClearMergedCellText(cell);
                    }
                }
            }
        }

        /// <summary>
        ///     Adds a new row to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the row should be inserted. If omitted, row is appended.</param>
        public void AddRow(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            int columns = Columns;
            A.TableRow row = new() { Height = 370840L };
            for (int c = 0; c < columns; c++) {
                A.TableCell cell = new(
                    new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                    new A.TableCellProperties()
                );
                row.Append(cell);
            }

            if (index.HasValue && index.Value < Rows) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(index.Value);
                table.InsertBefore(row, refRow);
            } else {
                table.Append(row);
            }
        }

        /// <summary>
        ///     Adds a new row cloned from a template row.
        /// </summary>
        public PowerPointTableRow AddRowFromTemplate(int templateRowIndex, int? index = null, bool clearText = true) {
            if (templateRowIndex < 0 || templateRowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(templateRowIndex));
            }

            A.Table table = TableElement;
            A.TableRow templateRow = table.Elements<A.TableRow>().ElementAt(templateRowIndex);
            A.TableRow newRow = (A.TableRow)templateRow.CloneNode(true);
            if (clearText) {
                foreach (A.TableCell cell in newRow.Elements<A.TableCell>()) {
                    ClearCellText(cell);
                }
            }

            int insertAt = index.HasValue ? Math.Min(index.Value, Rows) : Rows;
            if (insertAt < Rows) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(insertAt);
                table.InsertBefore(newRow, refRow);
            } else {
                table.Append(newRow);
            }

            return new PowerPointTableRow(this, newRow);
        }

        /// <summary>
        ///     Removes a row at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the row to remove.</param>
        public void RemoveRow(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableRow row = table.Elements<A.TableRow>().ElementAt(index);
            row.Remove();
        }

        /// <summary>
        ///     Adds a new column to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the column should be inserted. If omitted, column is appended.</param>
        public void AddColumn(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            A.GridColumn gridColumn = new() { Width = 3708400L };

            if (index.HasValue && index.Value < Columns) {
                A.GridColumn refCol = grid.Elements<A.GridColumn>().ElementAt(index.Value);
                grid.InsertBefore(gridColumn, refCol);
            } else {
                grid.Append(gridColumn);
            }

            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                A.TableCell cell = new(
                    new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                    new A.TableCellProperties()
                );

                if (index.HasValue && index.Value < Columns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(index.Value);
                    row.InsertBefore(cell, refCell);
                } else {
                    row.Append(cell);
                }
            }
        }

        /// <summary>
        ///     Adds a new column cloned from a template column.
        /// </summary>
        public PowerPointTableColumn AddColumnFromTemplate(int templateColumnIndex, int? index = null, bool clearText = true) {
            if (templateColumnIndex < 0 || templateColumnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(templateColumnIndex));
            }

            A.Table table = TableElement;
            A.TableGrid grid = table.TableGrid ?? throw new InvalidOperationException("Table grid is missing.");
            int existingColumns = Columns;
            A.GridColumn templateColumn = grid.Elements<A.GridColumn>().ElementAt(templateColumnIndex);
            A.GridColumn newColumn = (A.GridColumn)templateColumn.CloneNode(true);

            int insertAt = index.HasValue ? Math.Min(index.Value, existingColumns) : existingColumns;
            if (insertAt < existingColumns) {
                A.GridColumn refColumn = grid.Elements<A.GridColumn>().ElementAt(insertAt);
                grid.InsertBefore(newColumn, refColumn);
            } else {
                grid.Append(newColumn);
            }

            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                A.TableCell templateCell = row.Elements<A.TableCell>().ElementAt(templateColumnIndex);
                A.TableCell newCell = (A.TableCell)templateCell.CloneNode(true);
                if (clearText) {
                    ClearCellText(newCell);
                }

                if (insertAt < existingColumns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(insertAt);
                    row.InsertBefore(newCell, refCell);
                } else {
                    row.Append(newCell);
                }
            }

            return new PowerPointTableColumn(this, newColumn);
        }

        /// <summary>
        ///     Removes a column at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the column to remove.</param>
        public void RemoveColumn(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            grid.Elements<A.GridColumn>().ElementAt(index).Remove();
            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                row.Elements<A.TableCell>().ElementAt(index).Remove();
            }
        }

        /// <summary>
        ///     Binds data to the table, expanding rows/columns as needed.
        /// </summary>
        public void Bind<T>(IEnumerable<T> data, IEnumerable<PowerPointTableColumn<T>> columns,
            bool includeHeaders = true, int startRow = 0, int startColumn = 0) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }
            if (columns == null) {
                throw new ArgumentNullException(nameof(columns));
            }
            if (startRow < 0) {
                throw new ArgumentOutOfRangeException(nameof(startRow));
            }
            if (startColumn < 0) {
                throw new ArgumentOutOfRangeException(nameof(startColumn));
            }

            var items = data.ToList();
            var columnList = columns.ToList();
            if (columnList.Count == 0) {
                throw new ArgumentException("At least one column is required.", nameof(columns));
            }

            int requiredRows = items.Count + (includeHeaders ? 1 : 0);
            int requiredColumns = columnList.Count;

            while (Rows < startRow + requiredRows) {
                AddRow();
            }
            while (Columns < startColumn + requiredColumns) {
                AddColumn();
            }

            int rowIndex = startRow;
            if (includeHeaders) {
                for (int c = 0; c < columnList.Count; c++) {
                    GetCell(rowIndex, startColumn + c).Text = columnList[c].Header;
                }
                rowIndex++;
            }

            foreach (var item in items) {
                for (int c = 0; c < columnList.Count; c++) {
                    object? value = columnList[c].ValueSelector(item);
                    string text = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    GetCell(rowIndex, startColumn + c).Text = text;
                }
                rowIndex++;
            }
        }

        private static long ToEmus(double points) {
            return (long)Math.Round(points * EmusPerPoint);
        }

        private static void ClearCellText(A.TableCell cell) {
            if (cell.TextBody == null) {
                return;
            }

            cell.TextBody.RemoveAllChildren<A.Paragraph>();
            cell.TextBody.Append(new A.Paragraph(new A.Run(new A.Text(string.Empty))));
        }

        private static void ClearMergedCellText(A.TableCell cell) {
            ClearCellText(cell);
        }
    }
}
