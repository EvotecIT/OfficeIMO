namespace OfficeIMO.Pdf;

/// <summary>Fits the transformed appearance bounding box into an annotation rectangle.</summary>
internal static class PdfAppearancePlacement {
    internal static Matrix2D Read(PdfDictionary dictionary, Func<PdfObject, PdfObject?> resolve,
        double x, double y, double width, double height, out Matrix2D matrix) {
        matrix = Matrix2D.Identity;
        if (dictionary.Items.TryGetValue("Matrix", out var matrixValue)) {
            double[] values = ReadNumbers(matrixValue, 6, resolve);
            matrix = new Matrix2D(values[0], values[1], values[2], values[3], values[4], values[5]);
        }
        if (!dictionary.Items.TryGetValue("BBox", out var boxValue))
            throw new InvalidOperationException("The appearance has no bounding box.");
        double[] box = ReadNumbers(boxValue, 4, resolve);
        var p1 = matrix.Transform(box[0], box[1]);
        var p2 = matrix.Transform(box[0], box[3]);
        var p3 = matrix.Transform(box[2], box[1]);
        var p4 = matrix.Transform(box[2], box[3]);
        double minX = Math.Min(Math.Min(p1.X, p2.X), Math.Min(p3.X, p4.X));
        double minY = Math.Min(Math.Min(p1.Y, p2.Y), Math.Min(p3.Y, p4.Y));
        double maxX = Math.Max(Math.Max(p1.X, p2.X), Math.Max(p3.X, p4.X));
        double maxY = Math.Max(Math.Max(p1.Y, p2.Y), Math.Max(p3.Y, p4.Y));
        double sx = width / (maxX - minX), sy = height / (maxY - minY);
        double tx = x - minX * sx, ty = y - minY * sy;
        if (width <= 0 || height <= 0 || maxX <= minX || maxY <= minY || sx <= 0 || sy <= 0 ||
            !Finite(sx) || !Finite(sy) || !Finite(tx) || !Finite(ty))
            throw new InvalidOperationException("The appearance cannot be placed in a finite, nonempty rectangle.");
        // The Do operator applies the appearance's own Matrix after this outer transform.
        return new Matrix2D(sx, 0, 0, sy, tx, ty);
    }

    private static double[] ReadNumbers(PdfObject value, int count, Func<PdfObject, PdfObject?> resolve) {
        if (resolve(value) is not PdfArray array || array.Items.Count != count)
            throw new InvalidOperationException("The appearance geometry array is invalid.");
        var values = new double[count];
        for (int i = 0; i < count; i++) {
            if (resolve(array.Items[i]) is not PdfNumber number || !Finite(number.Value))
                throw new InvalidOperationException("The appearance geometry must contain finite numbers.");
            values[i] = number.Value;
        }
        return values;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
