namespace OfficeIMO.Pdf;

/// <summary>Fits the transformed appearance bounding box into an annotation rectangle.</summary>
internal static class PdfAppearancePlacement {
    /// <summary>
    /// Returns false when the appearance or the rectangle encloses no area. Such an appearance is clipped to its empty
    /// bounding box and paints nothing; PDF producers such as macOS PDFKit write them for empty widget states.
    /// Malformed geometry still throws.
    /// </summary>
    internal static bool TryRead(PdfDictionary dictionary, Func<PdfObject, PdfObject?> resolve,
        double x, double y, double width, double height, out Matrix2D placement, out Matrix2D matrix) {
        placement = Matrix2D.Identity;
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
        if (!Finite(x) || !Finite(y) || !Finite(width) || !Finite(height) ||
            !Finite(minX) || !Finite(minY) || !Finite(maxX) || !Finite(maxY))
            throw new InvalidOperationException("The transformed appearance geometry is not finite.");
        if (width <= 0 || height <= 0 || maxX <= minX || maxY <= minY) return false;
        double extentX = maxX - minX, extentY = maxY - minY;
        if (!Finite(extentX) || !Finite(extentY))
            throw new InvalidOperationException("The transformed appearance extents are not finite.");
        double sx = width / extentX, sy = height / extentY;
        double tx = x - minX * sx, ty = y - minY * sy;
        if (sx <= 0 || sy <= 0 || !Finite(sx) || !Finite(sy) || !Finite(tx) || !Finite(ty))
            throw new InvalidOperationException("The appearance placement cannot be represented by finite positive scales.");
        // The Do operator applies the appearance's own Matrix after this outer transform.
        placement = new Matrix2D(sx, 0, 0, sy, tx, ty);
        return true;
    }

    /// <summary>
    /// Identifies a proven no-paint area without requiring preserved opaque appearances to have usable placement.
    /// Missing or malformed geometry is not evidence of emptiness; its content still requires bounded admission.
    /// </summary>
    internal static bool IsKnownEmpty(PdfDictionary dictionary, Func<PdfObject, PdfObject?> resolve,
        double width, double height) {
        if (!dictionary.Items.ContainsKey("BBox")) return false;
        try {
            return !TryRead(dictionary, resolve, 0D, 0D, width, height, out _, out _);
        } catch (InvalidOperationException) {
            // The consuming mutation or renderer owns placement validation. Do not skip stream admission.
            return false;
        }
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
