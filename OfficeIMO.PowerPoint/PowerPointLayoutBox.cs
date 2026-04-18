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
        ///     Right position in centimeters.
        /// </summary>
        public double RightCm => PowerPointUnits.ToCentimeters(Right);

        /// <summary>
        ///     Bottom position in centimeters.
        /// </summary>
        public double BottomCm => PowerPointUnits.ToCentimeters(Bottom);

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
        ///     Right position in inches.
        /// </summary>
        public double RightInches => PowerPointUnits.ToInches(Right);

        /// <summary>
        ///     Bottom position in inches.
        /// </summary>
        public double BottomInches => PowerPointUnits.ToInches(Bottom);

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
        ///     Right position in points.
        /// </summary>
        public double RightPoints => PowerPointUnits.ToPoints(Right);

        /// <summary>
        ///     Bottom position in points.
        /// </summary>
        public double BottomPoints => PowerPointUnits.ToPoints(Bottom);

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
        ///     Splits the layout box into an equal row/column grid (EMU units).
        /// </summary>
        public PowerPointLayoutBox[,] SplitGrid(int rowCount, int columnCount, long rowGutterEmus, long columnGutterEmus) {
            if (rowCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowCount));
            }
            if (columnCount <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnCount));
            }
            if (rowGutterEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(rowGutterEmus));
            }
            if (columnGutterEmus < 0) {
                throw new ArgumentOutOfRangeException(nameof(columnGutterEmus));
            }

            PowerPointLayoutBox[,] grid = new PowerPointLayoutBox[rowCount, columnCount];
            PowerPointLayoutBox[] rows = SplitRows(rowCount, rowGutterEmus);
            for (int row = 0; row < rowCount; row++) {
                PowerPointLayoutBox[] columns = rows[row].SplitColumns(columnCount, columnGutterEmus);
                for (int column = 0; column < columnCount; column++) {
                    grid[row, column] = columns[column];
                }
            }

            return grid;
        }

        /// <summary>
        ///     Splits the layout box into an equal row/column grid in centimeters.
        /// </summary>
        public PowerPointLayoutBox[,] SplitGridCm(int rowCount, int columnCount, double rowGutterCm, double columnGutterCm) {
            return SplitGrid(rowCount, columnCount,
                PowerPointUnits.FromCentimeters(rowGutterCm),
                PowerPointUnits.FromCentimeters(columnGutterCm));
        }

        /// <summary>
        ///     Splits the layout box into an equal row/column grid in inches.
        /// </summary>
        public PowerPointLayoutBox[,] SplitGridInches(int rowCount, int columnCount, double rowGutterInches,
            double columnGutterInches) {
            return SplitGrid(rowCount, columnCount,
                PowerPointUnits.FromInches(rowGutterInches),
                PowerPointUnits.FromInches(columnGutterInches));
        }

        /// <summary>
        ///     Splits the layout box into an equal row/column grid in points.
        /// </summary>
        public PowerPointLayoutBox[,] SplitGridPoints(int rowCount, int columnCount, double rowGutterPoints,
            double columnGutterPoints) {
            return SplitGrid(rowCount, columnCount,
                PowerPointUnits.FromPoints(rowGutterPoints),
                PowerPointUnits.FromPoints(columnGutterPoints));
        }

        /// <summary>
        ///     Returns a new layout box inset equally on every side.
        /// </summary>
        public PowerPointLayoutBox InsetCm(double insetCm) {
            return InsetCm(insetCm, insetCm, insetCm, insetCm);
        }

        /// <summary>
        ///     Returns a new layout box inset horizontally and vertically.
        /// </summary>
        public PowerPointLayoutBox InsetCm(double horizontalCm, double verticalCm) {
            return InsetCm(horizontalCm, verticalCm, horizontalCm, verticalCm);
        }

        /// <summary>
        ///     Returns a new layout box inset by individual side values.
        /// </summary>
        public PowerPointLayoutBox InsetCm(double leftCm, double topCm, double rightCm, double bottomCm) {
            if (leftCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(leftCm));
            }
            if (topCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(topCm));
            }
            if (rightCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(rightCm));
            }
            if (bottomCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(bottomCm));
            }

            long left = PowerPointUnits.FromCentimeters(leftCm);
            long top = PowerPointUnits.FromCentimeters(topCm);
            long right = PowerPointUnits.FromCentimeters(rightCm);
            long bottom = PowerPointUnits.FromCentimeters(bottomCm);
            if (left + right >= Width || top + bottom >= Height) {
                throw new ArgumentOutOfRangeException(nameof(leftCm), "Insets exceed the layout box.");
            }

            return new PowerPointLayoutBox(Left + left, Top + top, Width - left - right, Height - top - bottom);
        }

        /// <summary>
        ///     Returns the top part of the layout box with the requested height.
        /// </summary>
        public PowerPointLayoutBox TakeTopCm(double heightCm) {
            if (heightCm <= 0) {
                throw new ArgumentOutOfRangeException(nameof(heightCm));
            }

            long height = PowerPointUnits.FromCentimeters(heightCm);
            if (height > Height) {
                throw new ArgumentOutOfRangeException(nameof(heightCm), "Requested height exceeds the layout box.");
            }

            return new PowerPointLayoutBox(Left, Top, Width, height);
        }

        /// <summary>
        ///     Returns the bottom part of the layout box with the requested height.
        /// </summary>
        public PowerPointLayoutBox TakeBottomCm(double heightCm) {
            if (heightCm <= 0) {
                throw new ArgumentOutOfRangeException(nameof(heightCm));
            }

            long height = PowerPointUnits.FromCentimeters(heightCm);
            if (height > Height) {
                throw new ArgumentOutOfRangeException(nameof(heightCm), "Requested height exceeds the layout box.");
            }

            return new PowerPointLayoutBox(Left, Bottom - height, Width, height);
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
