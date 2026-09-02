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
    internal static (double Width, double Height) GetVisualSize(
        PdfPageBox pageBox,
        int rotationDegrees,
        double userUnit = 1D) {
        double scale = NormalizeUserUnit(userUnit);
        return rotationDegrees == 90 || rotationDegrees == 270
            ? (pageBox.Height * scale, pageBox.Width * scale)
            : (pageBox.Width * scale, pageBox.Height * scale);
    }

    internal static Matrix2D CreateTransform(PdfPageBox pageBox, int rotationDegrees, double userUnit = 1D) {
        Matrix2D cropOrigin = Matrix2D.Translation(-pageBox.Left, -pageBox.Bottom);
        Matrix2D rotation = rotationDegrees switch {
            90 => new Matrix2D(0D, 1D, -1D, 0D, pageBox.Height, 0D),
            180 => new Matrix2D(-1D, 0D, 0D, -1D, pageBox.Width, pageBox.Height),
            270 => new Matrix2D(0D, -1D, 1D, 0D, 0D, pageBox.Width),
            _ => Matrix2D.Identity
        };
        Matrix2D transform = Matrix2D.Multiply(rotation, cropOrigin);
        double scale = NormalizeUserUnit(userUnit);
        return scale == 1D
            ? transform
            : Matrix2D.Multiply(new Matrix2D(scale, 0D, 0D, scale, 0D, 0D), transform);
    }

    internal static PdfVisualBounds TransformBounds(
        PdfPageBox pageBox,
        int rotationDegrees,
        double left,
        double bottom,
        double right,
        double top,
        double userUnit = 1D) {
        Matrix2D transform = CreateTransform(pageBox, rotationDegrees, userUnit);
        (double Width, double Height) size = GetVisualSize(pageBox, rotationDegrees, userUnit);
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

    internal static PdfVisualBounds TransformVisualBoundsToUser(
        PdfPageBox pageBox,
        int rotationDegrees,
        double left,
        double top,
        double right,
        double bottom,
        double userUnit = 1D) {
        Matrix2D transform = CreateTransform(pageBox, rotationDegrees, userUnit);
        double determinant = (transform.A * transform.D) - (transform.B * transform.C);
        if (Math.Abs(determinant) <= 0.000000001D) {
            throw new InvalidOperationException("The page visual transform is not invertible.");
        }

        (double Width, double Height) size = GetVisualSize(pageBox, rotationDegrees, userUnit);
        Matrix2D inverse = new Matrix2D(
            transform.D / determinant,
            -transform.B / determinant,
            -transform.C / determinant,
            transform.A / determinant,
            ((transform.C * transform.F) - (transform.D * transform.E)) / determinant,
            ((transform.B * transform.E) - (transform.A * transform.F)) / determinant);
        (double X, double Y) topLeft = inverse.Transform(left, size.Height - top);
        (double X, double Y) topRight = inverse.Transform(right, size.Height - top);
        (double X, double Y) bottomLeft = inverse.Transform(left, size.Height - bottom);
        (double X, double Y) bottomRight = inverse.Transform(right, size.Height - bottom);
        return new PdfVisualBounds(
            Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomLeft.X, bottomRight.X)),
            Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomLeft.Y, bottomRight.Y)),
            Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomLeft.X, bottomRight.X)),
            Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomLeft.Y, bottomRight.Y)));
    }

    internal static (double X, double Y) TransformVisualPointToUser(
        PdfPageBox pageBox,
        int rotationDegrees,
        double x,
        double y,
        double userUnit = 1D) {
        Matrix2D transform = CreateTransform(pageBox, rotationDegrees, userUnit);
        double determinant = (transform.A * transform.D) - (transform.B * transform.C);
        if (Math.Abs(determinant) <= 0.000000001D) {
            throw new InvalidOperationException("The page visual transform is not invertible.");
        }

        (double Width, double Height) size = GetVisualSize(pageBox, rotationDegrees, userUnit);
        Matrix2D inverse = new Matrix2D(
            transform.D / determinant,
            -transform.B / determinant,
            -transform.C / determinant,
            transform.A / determinant,
            ((transform.C * transform.F) - (transform.D * transform.E)) / determinant,
            ((transform.B * transform.E) - (transform.A * transform.F)) / determinant);
        return inverse.Transform(x, size.Height - y);
    }

    private static double NormalizeUserUnit(double userUnit) =>
        userUnit > 0D && !double.IsNaN(userUnit) && !double.IsInfinity(userUnit) ? userUnit : 1D;
}
