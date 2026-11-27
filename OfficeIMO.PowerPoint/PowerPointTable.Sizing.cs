using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Sets the width of a specific column as a physical length.
        /// </summary>
        public PowerPointTable SetColumnWidth(int columnIndex, OfficeLength width) {
            long nativeWidth = width.ToEmus();
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, ColumnCount);
            NthChild(GridNode, "gridCol", columnIndex).SetInt64Attribute("w", nativeWidth);
            SyncFrameWidthToGrid();
            MarkEdited();
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
            long width = GridWidth();
            OfficeXmlNode extents = FrameExtents;
            if (!extents.IsNone && width > 0) {
                extents.SetInt64Attribute("cx", width);
                MarkEdited();
            }
        }

        internal void SyncFrameHeightToGrid() {
            long height = GridHeight();
            OfficeXmlNode extents = FrameExtents;
            if (!extents.IsNone && height > 0) {
                extents.SetInt64Attribute("cy", height);
                MarkEdited();
            }
        }

        // The frame's a:ext in its p:xfrm, or none.
        private OfficeXmlNode FrameExtents => XmlElement.Child(OfficeXmlNamespaces.PresentationML, "xfrm").Child(D, "ext");

        private long GridWidth() {
            long width = 0;
            foreach (OfficeXmlNode column in GridNode.Children(D, "gridCol")) width += column.Int64Attribute("w") ?? 0L;
            return width;
        }

        private long GridHeight() {
            long height = 0;
            foreach (OfficeXmlNode row in TableNode.Children(D, "tr")) height += row.Int64Attribute("h") ?? 0L;
            return height;
        }


        /// <summary>
        ///     Gets the width of a specific column as a physical length.
        /// </summary>
        public OfficeLength GetColumnWidth(int columnIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, ColumnCount);
            return OfficeLength.Emus(NthChild(GridNode, "gridCol", columnIndex).Int64Attribute("w") ?? 0L);
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
                totalWidth = GridWidth();
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
                OfficeXmlNode column = NthChild(GridNode, "gridCol", adjustIndex);
                column.SetInt64Attribute("w", (column.Int64Attribute("w") ?? 0) + (totalWidth - assigned));
            }
            SyncFrameWidthToGrid();
            MarkEdited();
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
            NthChild(TableNode, "tr", rowIndex).SetInt64Attribute("h", nativeHeight);
            SyncFrameHeightToGrid();
            MarkEdited();
            return this;
        }

        /// <summary>
        ///     Gets the height of a specific row as a physical length.
        /// </summary>
        public OfficeLength GetRowHeight(int rowIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, RowCount);
            return OfficeLength.Emus(NthChild(TableNode, "tr", rowIndex).Int64Attribute("h") ?? 0L);
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
                totalHeight = GridHeight();
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
                OfficeXmlNode row = NthChild(TableNode, "tr", adjustIndex);
                row.SetInt64Attribute("h", (row.Int64Attribute("h") ?? 0) + (totalHeight - assigned));
            }
            SyncFrameHeightToGrid();
            MarkEdited();
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
                totalHeight = GridHeight();
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
                OfficeXmlNode row = NthChild(TableNode, "tr", adjustIndex);
                row.SetInt64Attribute("h", (row.Int64Attribute("h") ?? 0) + (totalHeight - assigned));
            }
            SyncFrameHeightToGrid();
            MarkEdited();
            return this;
        }

        private static long ToEmus(double points) {
            return (long)Math.Round(points * EmusPerPoint);
        }
    }
}
