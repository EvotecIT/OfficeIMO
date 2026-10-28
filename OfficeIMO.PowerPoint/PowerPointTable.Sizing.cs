using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Sets the width of a specific column as a physical length.
        /// </summary>
        public void SetColumnWidth(int columnIndex, OfficeLength width) {
            long nativeWidth = width.ToEmus();
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            column.Width = nativeWidth;
        }

        /// <summary>
        ///     Gets the width of a specific column as a physical length.
        /// </summary>
        public OfficeLength GetColumnWidth(int columnIndex) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            return OfficeLength.Emus(column.Width?.Value ?? 0L);
        }
        /// <summary>Sets physical sizes for the first rows or columns supplied.</summary>
        public void SetColumnWidths(params OfficeLength[] widths) {
            if (widths == null) throw new ArgumentNullException(nameof(widths));
            int count = Math.Min(widths.Length, Columns);
            for (int i = 0; i < count; i++) {
                SetColumnWidth(i, widths[i]);
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
        ///     Sets the height of a specific row as a physical length.
        /// </summary>
        public void SetRowHeight(int rowIndex, OfficeLength height) {
            long nativeHeight = height.ToEmus();
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            row.Height = nativeHeight;
        }

        /// <summary>
        ///     Gets the height of a specific row as a physical length.
        /// </summary>
        public OfficeLength GetRowHeight(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            return OfficeLength.Emus(row.Height?.Value ?? 0L);
        }
        /// <summary>Sets physical sizes for the first rows or columns supplied.</summary>
        public void SetRowHeights(params OfficeLength[] heights) {
            if (heights == null) throw new ArgumentNullException(nameof(heights));
            int count = Math.Min(heights.Length, Rows);
            for (int i = 0; i < count; i++) {
                SetRowHeight(i, heights[i]);
            }
        }

        /// <summary>
        ///     Sets row heights evenly across the table height.
        /// </summary>
        public void SetRowHeightsEvenly() {
            if (Rows == 0) {
                return;
            }

            long totalHeight = Height.ToEmus();
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
                SetRowHeight(i, OfficeLength.Emus(rowHeight));
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
        }

        private static long ToEmus(double points) {
            return (long)Math.Round(points * EmusPerPoint);
        }
    }
}