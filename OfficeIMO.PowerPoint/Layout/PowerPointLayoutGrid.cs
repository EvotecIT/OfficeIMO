using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>A box in points.</summary>
    internal readonly struct PowerPointLayoutRect {
        internal PowerPointLayoutRect(double x, double y, double width, double height) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Right => X + Width;
        internal double Bottom => Y + Height;

        internal PowerPointLayoutRect WithY(double y) => new(X, y, Width, Height);
        internal PowerPointLayoutRect WithHeight(double height) => new(X, Y, Width, Math.Max(0D, height));
        internal PowerPointLayoutRect WithWidth(double width) => new(X, Y, Math.Max(0D, width), Height);

        internal PowerPointLayoutRect Inset(double left, double top, double right, double bottom) =>
            new(X + left, Y + top, Math.Max(0D, Width - left - right), Math.Max(0D, Height - top - bottom));

        public override string ToString() => $"{X:0.#},{Y:0.#} {Width:0.#}x{Height:0.#}";
    }

    /// <summary>
    ///     The 12-column, 12-row slide grid. Side and top margins come from the theme; the live area ends at a
    ///     fixed 488 pt (on a 540 pt slide) so the footer band never moves.
    /// </summary>
    internal sealed class PowerPointLayoutGrid {
        internal const int Columns = 12;
        internal const int Rows = 12;
        internal const double RowGutter = 12D;
        internal const double FooterOffsetFromBottom = 34D; // 506 pt on a 540 pt slide
        internal const double FooterHeight = 16D;
        internal const double LiveBottomOffset = 52D; // 488 pt on a 540 pt slide

        internal PowerPointLayoutGrid(double slideWidth, double slideHeight, double marginLeft, double marginRight,
            double marginTop) {
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;
            bool narrow = slideWidth / slideHeight < 1.5D;
            Gutter = narrow ? 18D : 24D;
            // 4:3 keeps the reference 42 pt side margins (powerpoint-layouts.json, Screen4x3).
            MarginLeft = narrow ? Math.Min(marginLeft, 42D) : marginLeft;
            MarginRight = narrow ? Math.Min(marginRight, 42D) : marginRight;
            MarginTop = marginTop;
            LiveBottom = slideHeight - LiveBottomOffset;
            Column = (slideWidth - MarginLeft - MarginRight - (Columns - 1) * Gutter) / Columns;
            Row = (LiveBottom - MarginTop - (Rows - 1) * RowGutter) / Rows;
            FooterY = slideHeight - FooterOffsetFromBottom;
            IsNarrow = narrow;
        }

        private PowerPointLayoutGrid(PowerPointLayoutGrid parent, PowerPointLayoutRect region) {
            SlideWidth = parent.SlideWidth;
            SlideHeight = parent.SlideHeight;
            Gutter = parent.Gutter;
            MarginLeft = region.X;
            MarginRight = parent.SlideWidth - region.Right;
            MarginTop = region.Y;
            LiveBottom = region.Bottom;
            Column = Math.Max(1D, (region.Width - (Columns - 1) * Gutter) / Columns);
            Row = parent.Row;
            FooterY = parent.FooterY;
            IsNarrow = parent.IsNarrow;
        }

        private PowerPointLayoutGrid(double slideWidth, double slideHeight, PowerPointLayoutRect body, double top) {
            SlideWidth = slideWidth;
            SlideHeight = slideHeight;
            Gutter = Math.Max(8D, body.Width * 0.025D);
            MarginLeft = body.X;
            MarginRight = slideWidth - body.Right;
            MarginTop = top;
            LiveBottom = body.Bottom;
            Column = (body.Width - (Columns - 1) * Gutter) / Columns;
            Row = Math.Max(8D, (LiveBottom - MarginTop - (Rows - 1) * RowGutter) / Rows);
            FooterY = slideHeight - FooterOffsetFromBottom;
            IsNarrow = slideWidth / slideHeight < 1.5D;
        }

        /// <summary>
        ///     The grid re-derived from a template's body placeholder: its left and right edges are the margins, the gutter
        ///     is 2.5 % of its width, the title top is the top margin and the body bottom ends the live area.
        /// </summary>
        internal static PowerPointLayoutGrid ForTemplate(double slideWidth, double slideHeight, PowerPointLayoutRect body, double top) =>
            new(slideWidth, slideHeight, body, top);

        /// <summary>
        ///     A 12-column grid over <paramref name="region" /> with the parent's gutter and row height, so a
        ///     composition placed in a content region keeps the deck's rhythm. Row 1 starts at the region top.
        /// </summary>
        internal static PowerPointLayoutGrid ForRegion(PowerPointLayoutGrid parent, PowerPointLayoutRect region) => new(parent, region);

        internal static PowerPointLayoutGrid For(PowerPointPresentation presentation, PowerPointLayoutTheme theme) {
            PowerPointSlideSize size = presentation.SlideSize;
            return new PowerPointLayoutGrid(PowerPointUnits.ToPoints(size.WidthEmus), PowerPointUnits.ToPoints(size.HeightEmus),
                theme.MarginLeft, theme.MarginRight, theme.MarginTop);
        }

        internal double SlideWidth { get; }
        internal double SlideHeight { get; }
        internal double Gutter { get; }
        internal double MarginLeft { get; }
        internal double MarginRight { get; }
        internal double MarginTop { get; }
        internal double LiveBottom { get; }
        internal double Column { get; }
        internal double Row { get; }
        internal double FooterY { get; }
        internal bool IsNarrow { get; }
        internal double LiveWidth => SlideWidth - MarginLeft - MarginRight;

        /// <summary>Left edge of 1-based column <paramref name="column" />.</summary>
        internal double X(int column) => MarginLeft + (column - 1) * (Column + Gutter);

        internal double W(int span) => span * Column + (span - 1) * Gutter;

        /// <summary>Top edge of 1-based row <paramref name="row" />.</summary>
        internal double Y(int row) => MarginTop + (row - 1) * (Row + RowGutter);

        internal double H(int rows) => rows * Row + (rows - 1) * RowGutter;

        internal PowerPointLayoutRect Box(int column, int span, int row, int rows) =>
            new(X(column), Y(row), W(span), H(rows));

        /// <summary>First row whose top is at least <paramref name="gap" /> below <paramref name="bottom" />.</summary>
        internal int FirstRowBelow(double bottom, double gap) {
            for (int row = 1; row <= Rows; row++) {
                if (Y(row) >= bottom + gap - 0.01D) return row;
            }

            return Rows;
        }
    }
}
