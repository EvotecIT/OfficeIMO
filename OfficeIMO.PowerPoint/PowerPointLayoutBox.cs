using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a rectangular layout region on a slide.
    /// </summary>
    public readonly struct PowerPointLayoutBox {
        /// <summary>
        ///     Creates a new layout box in EMUs.
        /// </summary>
        /// <param name="left">Left position in EMUs.</param>
        /// <param name="top">Top position in EMUs.</param>
        /// <param name="width">Width in EMUs.</param>
        /// <param name="height">Height in EMUs.</param>
        public PowerPointLayoutBox(long left, long top, long width, long height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        /// <summary>
        ///     Creates a new layout box in centimeters.
        /// </summary>
        public static PowerPointLayoutBox FromCentimeters(double leftCm, double topCm, double widthCm, double heightCm) {
            return new PowerPointLayoutBox(
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Creates a new layout box in inches.
        /// </summary>
        public static PowerPointLayoutBox FromInches(double leftInches, double topInches, double widthInches, double heightInches) {
            return new PowerPointLayoutBox(
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Creates a new layout box in points.
        /// </summary>
        public static PowerPointLayoutBox FromPoints(double leftPoints, double topPoints, double widthPoints, double heightPoints) {
            return new PowerPointLayoutBox(
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Left position in EMUs.
        /// </summary>
        public long Left { get; }

        /// <summary>
        ///     Top position in EMUs.
        /// </summary>
        public long Top { get; }

        /// <summary>
        ///     Width in EMUs.
        /// </summary>
        public long Width { get; }

        /// <summary>
        ///     Height in EMUs.
        /// </summary>
        public long Height { get; }

        /// <summary>
        ///     Right position in EMUs.
        /// </summary>
        public long Right => Left + Width;

        /// <summary>
        ///     Bottom position in EMUs.
        /// </summary>
        public long Bottom => Top + Height;

        /// <summary>
        ///     Left position in centimeters.
        /// </summary>
        public double LeftCm => PowerPointUnits.ToCentimeters(Left);

        /// <summary>
        ///     Top position in centimeters.
        /// </summary>
        public double TopCm => PowerPointUnits.ToCentimeters(Top);

        /// <summary>
        ///     Width in centimeters.
        /// </summary>
        public double WidthCm => PowerPointUnits.ToCentimeters(Width);

        /// <summary>
        ///     Height in centimeters.
        /// </summary>
        public double HeightCm => PowerPointUnits.ToCentimeters(Height);

        /// <summary>
        ///     Left position in inches.
        /// </summary>
        public double LeftInches => PowerPointUnits.ToInches(Left);

        /// <summary>
        ///     Top position in inches.
        /// </summary>
        public double TopInches => PowerPointUnits.ToInches(Top);

        /// <summary>
        ///     Width in inches.
        /// </summary>
        public double WidthInches => PowerPointUnits.ToInches(Width);

        /// <summary>
        ///     Height in inches.
        /// </summary>
        public double HeightInches => PowerPointUnits.ToInches(Height);

        /// <summary>
        ///     Left position in points.
        /// </summary>
        public double LeftPoints => PowerPointUnits.ToPoints(Left);

        /// <summary>
        ///     Top position in points.
        /// </summary>
        public double TopPoints => PowerPointUnits.ToPoints(Top);

        /// <summary>
        ///     Width in points.
        /// </summary>
        public double WidthPoints => PowerPointUnits.ToPoints(Width);

        /// <summary>
        ///     Height in points.
        /// </summary>
        public double HeightPoints => PowerPointUnits.ToPoints(Height);

        /// <summary>
        ///     Splits the layout box into equal columns (EMU units).
        /// </summary>
        public PowerPointLayoutBox[] SplitColumns(int columnCount, long gutterEmus) {
            if (columnCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnCount));
            }
            if (gutterEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus));
            }

            long totalGutter = gutterEmus * (columnCount - 1);
            if (totalGutter > Width) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus), "Gutter exceeds available width.");
            }

            long columnWidth = (Width - totalGutter) / columnCount;
            if (columnWidth <= 0) {
                throw new InvalidOperationException("Column width is not positive.");
            }

            var columns = new PowerPointLayoutBox[columnCount];
            long left = Left;
            for (int i = 0; i < columnCount; i++) {
                columns[i] = new PowerPointLayoutBox(left, Top, columnWidth, Height);
                left += columnWidth + gutterEmus;
            }

            return columns;
        }

        /// <summary>
        ///     Splits the layout box into equal columns in centimeters.
        /// </summary>
        public PowerPointLayoutBox[] SplitColumnsCm(int columnCount, double gutterCm) {
            return SplitColumns(columnCount, PowerPointUnits.FromCentimeters(gutterCm));
        }

        /// <summary>
        ///     Splits the layout box into equal columns in inches.
        /// </summary>
        public PowerPointLayoutBox[] SplitColumnsInches(int columnCount, double gutterInches) {
            return SplitColumns(columnCount, PowerPointUnits.FromInches(gutterInches));
        }

        /// <summary>
        ///     Splits the layout box into equal columns in points.
        /// </summary>
        public PowerPointLayoutBox[] SplitColumnsPoints(int columnCount, double gutterPoints) {
            return SplitColumns(columnCount, PowerPointUnits.FromPoints(gutterPoints));
        }

        /// <summary>
        ///     Splits the layout box into equal rows (EMU units).
        /// </summary>
        public PowerPointLayoutBox[] SplitRows(int rowCount, long gutterEmus) {
            if (rowCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowCount));
            }
            if (gutterEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus));
            }

            long totalGutter = gutterEmus * (rowCount - 1);
            if (totalGutter > Height) {
                throw new ArgumentOutOfRangeException(nameof(gutterEmus), "Gutter exceeds available height.");
            }

            long rowHeight = (Height - totalGutter) / rowCount;
            if (rowHeight <= 0) {
                throw new InvalidOperationException("Row height is not positive.");
            }

            var rows = new PowerPointLayoutBox[rowCount];
            long top = Top;
            for (int i = 0; i < rowCount; i++) {
                rows[i] = new PowerPointLayoutBox(Left, top, Width, rowHeight);
                top += rowHeight + gutterEmus;
            }

            return rows;
        }

        /// <summary>
        ///     Splits the layout box into equal rows in centimeters.
        /// </summary>
        public PowerPointLayoutBox[] SplitRowsCm(int rowCount, double gutterCm) {
            return SplitRows(rowCount, PowerPointUnits.FromCentimeters(gutterCm));
        }

        /// <summary>
        ///     Splits the layout box into equal rows in inches.
        /// </summary>
        public PowerPointLayoutBox[] SplitRowsInches(int rowCount, double gutterInches) {
            return SplitRows(rowCount, PowerPointUnits.FromInches(gutterInches));
        }

        /// <summary>
        ///     Splits the layout box into equal rows in points.
        /// </summary>
        public PowerPointLayoutBox[] SplitRowsPoints(int rowCount, double gutterPoints) {
            return SplitRows(rowCount, PowerPointUnits.FromPoints(gutterPoints));
        }

        /// <summary>
        ///     Applies the layout box to the provided shape.
        /// </summary>
        /// <param name="shape">Shape to update.</param>
        public void ApplyTo(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.Left = Left;
            shape.Top = Top;
            shape.Width = Width;
            shape.Height = Height;
        }
    }
}
