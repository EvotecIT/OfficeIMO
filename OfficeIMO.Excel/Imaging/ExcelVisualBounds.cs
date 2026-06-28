namespace OfficeIMO.Excel {
    internal readonly struct ExcelVisualBounds {
        internal ExcelVisualBounds(double x, double y, double width, double height) {
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

        internal bool IsEmpty => Width <= 0D || Height <= 0D;

        internal double IntersectionArea(ExcelVisualBounds other) {
            double left = Math.Max(X, other.X);
            double top = Math.Max(Y, other.Y);
            double right = Math.Min(Right, other.Right);
            double bottom = Math.Min(Bottom, other.Bottom);
            double width = right - left;
            double height = bottom - top;
            return width <= 0D || height <= 0D ? 0D : width * height;
        }
    }
}
