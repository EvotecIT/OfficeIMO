using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>Represents a rectangular layout region on a slide.</summary>
    public readonly struct PowerPointLayoutBox {
        /// <summary>Creates a region with explicit position and size.</summary>
        public PowerPointLayoutBox(OfficeLength left, OfficeLength top, OfficeLength width, OfficeLength height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        /// <summary>Left position of the region.</summary>
        public OfficeLength Left { get; }
        /// <summary>Top position of the region.</summary>
        public OfficeLength Top { get; }
        /// <summary>Width of the region.</summary>
        public OfficeLength Width { get; }
        /// <summary>Height of the region.</summary>
        public OfficeLength Height { get; }
        /// <summary>Right position of the region.</summary>
        public OfficeLength Right => Left + Width;
        /// <summary>Bottom position of the region.</summary>
        public OfficeLength Bottom => Top + Height;

        /// <summary>Splits the region into equal columns separated by a gutter.</summary>
        public PowerPointLayoutBox[] SplitColumns(int columnCount, OfficeLength gutter) {
            if (columnCount <= 0) throw new ArgumentOutOfRangeException(nameof(columnCount));
            if (gutter < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(gutter));
            long totalGutter = checked(gutter.ToEmus() * (columnCount - 1));
            if (totalGutter > Width.ToEmus()) throw new ArgumentOutOfRangeException(nameof(gutter), "Gutter exceeds available width.");
            var columnWidth = OfficeLength.Emus(checked(Width.ToEmus() - totalGutter) / columnCount);
            if (columnWidth <= OfficeLength.Zero) throw new InvalidOperationException("Column width is not positive.");
            var columns = new PowerPointLayoutBox[columnCount];
            OfficeLength left = Left;
            for (int i = 0; i < columnCount; i++) {
                columns[i] = new PowerPointLayoutBox(left, Top, columnWidth, Height);
                left += columnWidth + gutter;
            }
            return columns;
        }

        /// <summary>Splits the region into equal rows separated by a gutter.</summary>
        public PowerPointLayoutBox[] SplitRows(int rowCount, OfficeLength gutter) {
            if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));
            if (gutter < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(gutter));
            long totalGutter = checked(gutter.ToEmus() * (rowCount - 1));
            if (totalGutter > Height.ToEmus()) throw new ArgumentOutOfRangeException(nameof(gutter), "Gutter exceeds available height.");
            var rowHeight = OfficeLength.Emus(checked(Height.ToEmus() - totalGutter) / rowCount);
            if (rowHeight <= OfficeLength.Zero) throw new InvalidOperationException("Row height is not positive.");
            var rows = new PowerPointLayoutBox[rowCount];
            OfficeLength top = Top;
            for (int i = 0; i < rowCount; i++) {
                rows[i] = new PowerPointLayoutBox(Left, top, Width, rowHeight);
                top += rowHeight + gutter;
            }
            return rows;
        }

        /// <summary>Splits the region into an equal row and column grid.</summary>
        public PowerPointLayoutBox[,] SplitGrid(int rowCount, int columnCount, OfficeLength rowGutter, OfficeLength columnGutter) {
            if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));
            if (columnCount <= 0) throw new ArgumentOutOfRangeException(nameof(columnCount));
            if (rowGutter < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(rowGutter));
            if (columnGutter < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(columnGutter));
            var grid = new PowerPointLayoutBox[rowCount, columnCount];
            var rows = SplitRows(rowCount, rowGutter);
            for (int row = 0; row < rowCount; row++) {
                var columns = rows[row].SplitColumns(columnCount, columnGutter);
                for (int column = 0; column < columnCount; column++) grid[row, column] = columns[column];
            }
            return grid;
        }

        /// <summary>Returns a region inset equally on every side.</summary>
        public PowerPointLayoutBox Inset(OfficeLength inset) => Inset(inset, inset, inset, inset);
        /// <summary>Returns a region inset horizontally and vertically.</summary>
        public PowerPointLayoutBox Inset(OfficeLength horizontal, OfficeLength vertical) => Inset(horizontal, vertical, horizontal, vertical);
        /// <summary>Returns a region inset by individual side values.</summary>
        public PowerPointLayoutBox Inset(OfficeLength left, OfficeLength top, OfficeLength right, OfficeLength bottom) {
            if (left < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(left));
            if (top < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(top));
            if (right < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(right));
            if (bottom < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(bottom));
            if (left + right >= Width || top + bottom >= Height) throw new ArgumentOutOfRangeException(nameof(left), "Insets exceed the layout box.");
            return new PowerPointLayoutBox(Left + left, Top + top, Width - left - right, Height - top - bottom);
        }

        /// <summary>Returns the top part with the requested height.</summary>
        public PowerPointLayoutBox TakeTop(OfficeLength height) {
            ValidateHeight(height);
            return new PowerPointLayoutBox(Left, Top, Width, height);
        }
        /// <summary>Returns the bottom part with the requested height.</summary>
        public PowerPointLayoutBox TakeBottom(OfficeLength height) {
            ValidateHeight(height);
            return new PowerPointLayoutBox(Left, Bottom - height, Width, height);
        }
        private void ValidateHeight(OfficeLength height) {
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            if (height > Height) throw new ArgumentOutOfRangeException(nameof(height), "Requested height exceeds the layout box.");
        }

        /// <summary>Applies the region's position and size to a shape.</summary>
        public void ApplyTo(PowerPointShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            shape.Bounds = this;
        }
    }
}
