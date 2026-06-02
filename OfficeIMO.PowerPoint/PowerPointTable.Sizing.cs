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

        private static long ToEmus(double points) {
            return (long)Math.Round(points * EmusPerPoint);
        }
    }
}
