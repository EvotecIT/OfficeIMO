namespace OfficeIMO.Pdf;

internal readonly struct PdfVisualBounds {
    internal PdfVisualBounds(double left, double top, double right, double bottom) {
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    internal double Left { get; }
    internal double Top { get; }
    internal double Right { get; }
    internal double Bottom { get; }
    internal double Width => Right - Left;
    internal double Height => Bottom - Top;
}

internal static class PdfVisualCoordinateMapper {
    internal static (double Width, double Height) GetVisualSize(PdfPageBox pageBox, int rotationDegrees) =>
        rotationDegrees == 90 || rotationDegrees == 270
            ? (pageBox.Height, pageBox.Width)
            : (pageBox.Width, pageBox.Height);

    internal static Matrix2D CreateTransform(PdfPageBox pageBox, int rotationDegrees) {
        Matrix2D cropOrigin = Matrix2D.Translation(-pageBox.Left, -pageBox.Bottom);
        Matrix2D rotation = rotationDegrees switch {
            90 => new Matrix2D(0D, 1D, -1D, 0D, pageBox.Height, 0D),
            180 => new Matrix2D(-1D, 0D, 0D, -1D, pageBox.Width, pageBox.Height),
            270 => new Matrix2D(0D, -1D, 1D, 0D, 0D, pageBox.Width),
            _ => Matrix2D.Identity
        };
        return Matrix2D.Multiply(rotation, cropOrigin);
    }

    internal static PdfVisualBounds TransformBounds(
        PdfPageBox pageBox,
        int rotationDegrees,
        double left,
        double bottom,
        double right,
        double top) {
        Matrix2D transform = CreateTransform(pageBox, rotationDegrees);
        (double Width, double Height) size = GetVisualSize(pageBox, rotationDegrees);
        (double X, double Y) bottomLeft = transform.Transform(left, bottom);
        (double X, double Y) topLeft = transform.Transform(left, top);
        (double X, double Y) bottomRight = transform.Transform(right, bottom);
        (double X, double Y) topRight = transform.Transform(right, top);
        double visualLeft = Math.Min(Math.Min(bottomLeft.X, topLeft.X), Math.Min(bottomRight.X, topRight.X));
        double visualRight = Math.Max(Math.Max(bottomLeft.X, topLeft.X), Math.Max(bottomRight.X, topRight.X));
        double visualBottom = Math.Min(Math.Min(bottomLeft.Y, topLeft.Y), Math.Min(bottomRight.Y, topRight.Y));
        double visualTop = Math.Max(Math.Max(bottomLeft.Y, topLeft.Y), Math.Max(bottomRight.Y, topRight.Y));
        return new PdfVisualBounds(
            visualLeft,
            size.Height - visualTop,
            visualRight,
            size.Height - visualBottom);
    }
}
