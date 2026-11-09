using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Sets the width of a specific column as a physical length.
        /// </summary>
        public PowerPointTable SetColumnWidth(int columnIndex, OfficeLength width) {
            long nativeWidth = width.ToEmus();
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, ColumnCount);
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            column.Width = nativeWidth;
            SyncFrameWidthToGrid();
            return this;
        }

        /// <summary>
        ///     Keeps the graphic frame equal to the table grid (sum of column widths and row heights). Office applications
        ///     draw the grid, so a frame that disagrees misplaces selection handles and makes
        ///     frame-based renderers stretch the table.
        /// </summary>
        internal void SyncFrameToGrid() {
            SyncFrameWidthToGrid();
            SyncFrameHeightToGrid();
        }

        internal void SyncFrameWidthToGrid() {
            long width = TableElement.TableGrid?.Elements<A.GridColumn>().Sum(c => c.Width?.Value ?? 0L) ?? 0L;
            A.Extents? extents = Frame.Transform?.Extents;
            if (extents != null && width > 0) extents.Cx = width;
        }

        internal void SyncFrameHeightToGrid() {
            long height = TableElement.Elements<A.TableRow>().Sum(r => r.Height?.Value ?? 0L);
            A.Extents? extents = Frame.Transform?.Extents;
            if (extents != null && height > 0) extents.Cy = height;
        }


        /// <summary>
        ///     Gets the width of a specific column as a physical length.
        /// </summary>
        public OfficeLength GetColumnWidth(int columnIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, ColumnCount);
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            return OfficeLength.Emus(column.Width?.Value ?? 0L);
        }
        /// <summary>Sets physical sizes for the first rows or columns supplied.</summary>
        public PowerPointTable SetColumnWidths(params OfficeLength[] widths) {
            ArgumentNullException.ThrowIfNull(widths);
            int count = Math.Min(widths.Length, ColumnCount);
            for (int i = 0; i < count; i++) {
                SetColumnWidth(i, widths[i]);
            }
            return this;
        }

        /// <summary>
        ///     Sets column widths proportionally using ratios.
        /// </summary>
        public PowerPointTable SetColumnWidthsByRatio(params double[] ratios) {
            ArgumentNullException.ThrowIfNull(ratios);
            if (ratios.Length == 0) {
                throw new ArgumentException("At least one ratio is required.", nameof(ratios));
            }

            int count = Math.Min(ratios.Length, ColumnCount);
            double totalRatio = 0;
            for (int i = 0; i < count; i++) {
                if (ratios[i] <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(ratios), "Ratios must be positive.");
                }
                totalRatio += ratios[i];
            }

            long totalWidth = Width.ToEmus();
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
                SetColumnWidth(i, OfficeLength.Emus(width));
                assigned += width;
            }

            if (assigned != totalWidth && count > 0) {
                int adjustIndex = count - 1;
                A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(adjustIndex);
                column.Width = (column.Width ?? 0) + (totalWidth - assigned);
            }
            SyncFrameWidthToGrid();
            return this;
        }

        /// <summary>
        ///     Sets column widths evenly across the table width.
        /// </summary>
        public PowerPointTable SetColumnWidthsEvenly() {
            if (ColumnCount == 0) {
                return this;
            }
            SetColumnWidthsByRatio(Enumerable.Repeat(1d, ColumnCount).ToArray());
            return this;
        }

        /// <summary>
        ///     Sets the height of a specific row as a physical length.
        /// </summary>
        public PowerPointTable SetRowHeight(int rowIndex, OfficeLength height) {
            long nativeHeight = height.ToEmus();
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, RowCount);
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            row.Height = nativeHeight;
            SyncFrameHeightToGrid();
            return this;
        }

        /// <summary>
        ///     Gets the height of a specific row as a physical length.
        /// </summary>
        public OfficeLength GetRowHeight(int rowIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, RowCount);
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            return OfficeLength.Emus(row.Height?.Value ?? 0L);
        }
        /// <summary>Sets physical sizes for the first rows or columns supplied.</summary>
        public PowerPointTable SetRowHeights(params OfficeLength[] heights) {
            ArgumentNullException.ThrowIfNull(heights);
            int count = Math.Min(heights.Length, RowCount);
            for (int i = 0; i < count; i++) {
                SetRowHeight(i, heights[i]);
            }
            return this;
        }

        /// <summary>
        ///     Sets row heights evenly across the table height.
        /// </summary>
        public PowerPointTable SetRowHeightsEvenly() {
            if (RowCount == 0) {
                return this;
            }

            long totalHeight = Height.ToEmus();
            if (totalHeight <= 0) {
                totalHeight = TableElement.Elements<A.TableRow>()
                    .Sum(r => r.Height?.Value ?? 0L);
            }
            if (totalHeight <= 0) {
                return this;
            }

            long rowHeight = (long)Math.Floor(totalHeight / (double)RowCount);
            long assigned = 0;
            for (int i = 0; i < RowCount; i++) {
                SetRowHeight(i, OfficeLength.Emus(rowHeight));
                assigned += rowHeight;
            }

            if (assigned != totalHeight && RowCount > 0) {
                int adjustIndex = RowCount - 1;
                A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(adjustIndex);
                row.Height = (row.Height ?? 0) + (totalHeight - assigned);
            }
            SyncFrameHeightToGrid();
            return this;
        }

        /// <summary>
        ///     Sets row heights proportionally using ratios.
        /// </summary>
        public PowerPointTable SetRowHeightsByRatio(params double[] ratios) {
            ArgumentNullException.ThrowIfNull(ratios);
            if (ratios.Length == 0) {
                throw new ArgumentException("At least one ratio is required.", nameof(ratios));
            }

            int count = Math.Min(ratios.Length, RowCount);
            double totalRatio = 0;
            for (int i = 0; i < count; i++) {
                if (ratios[i] <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(ratios), "Ratios must be positive.");
                }
                totalRatio += ratios[i];
            }

            long totalHeight = Height.ToEmus();
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
                SetRowHeight(i, OfficeLength.Emus(height));
                assigned += height;
            }

            if (assigned != totalHeight && count > 0) {
                int adjustIndex = count - 1;
                A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(adjustIndex);
                row.Height = (row.Height ?? 0) + (totalHeight - assigned);
            }
            SyncFrameHeightToGrid();
            return this;
        }

        private static long ToEmus(double points) {
            return (long)Math.Round(points * EmusPerPoint);
        }
    }
}
