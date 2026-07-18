namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationFlattener {
    private static string FormatNumber(double value) {
        if (Math.Abs(value % 1D) < 0.0000001D) {
            return ((long)Math.Round(value)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static int FindCatalogObjectNumber(Dictionary<int, PdfIndirectObject> objects, string? trailerRaw) {
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is null) {
            return 0;
        }

        foreach (var entry in objects) {
            if (ReferenceEquals(entry.Value.Value, catalog)) {
                return entry.Key;
            }
        }

        return 0;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return PdfObjectLookup.Resolve(objects, value);
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return ResolveObject(objects, value) as PdfDictionary;
    }

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static bool TryReadRectCoordinates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, out double x, out double y, out double width, out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (!dictionary.Items.TryGetValue("Rect", out var rectObject) ||
            ResolveObject(objects, rectObject) is not PdfArray rect ||
            rect.Items.Count < 4 ||
            ResolveObject(objects, rect.Items[0]) is not PdfNumber x1 ||
            ResolveObject(objects, rect.Items[1]) is not PdfNumber y1 ||
            ResolveObject(objects, rect.Items[2]) is not PdfNumber x2 ||
            ResolveObject(objects, rect.Items[3]) is not PdfNumber y2) {
            return false;
        }

        x = Math.Min(x1.Value, x2.Value);
        y = Math.Min(y1.Value, y2.Value);
        width = Math.Abs(x2.Value - x1.Value);
        height = Math.Abs(y2.Value - y1.Value);
        return width > 0D && height > 0D;
    }

    private static AppearancePlacement ReadAppearancePlacement(Dictionary<int, PdfIndirectObject> objects, PdfReference appearanceReference, double x, double y, double width, double height) {
        if (!PdfObjectLookup.TryGet(objects, appearanceReference, out var appearanceObject) ||
            appearanceObject.Value is not PdfStream appearanceStream ||
            !TryReadBoxCoordinates(objects, appearanceStream.Dictionary, "BBox", out double bboxX, out double bboxY, out double bboxWidth, out double bboxHeight)) {
            return new AppearancePlacement(1D, 0D, 0D, 1D, x, y);
        }

        double scaleX = width / bboxWidth;
        double scaleY = height / bboxHeight;
        var placement = new AppearancePlacement(scaleX, 0D, 0D, scaleY, x - bboxX * scaleX, y - bboxY * scaleY);
        return TryReadMatrix(objects, appearanceStream.Dictionary, out AppearancePlacement matrix) &&
            TryInvertMatrix(matrix, out AppearancePlacement inverseMatrix)
            ? MultiplyMatrices(placement, inverseMatrix)
            : placement;
    }

    private static bool TryReadMatrix(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, out AppearancePlacement matrix) {
        matrix = new AppearancePlacement(1D, 0D, 0D, 1D, 0D, 0D);
        if (!dictionary.Items.TryGetValue("Matrix", out var matrixObject) ||
            ResolveObject(objects, matrixObject) is not PdfArray matrixArray ||
            matrixArray.Items.Count < 6 ||
            ResolveObject(objects, matrixArray.Items[0]) is not PdfNumber a ||
            ResolveObject(objects, matrixArray.Items[1]) is not PdfNumber b ||
            ResolveObject(objects, matrixArray.Items[2]) is not PdfNumber c ||
            ResolveObject(objects, matrixArray.Items[3]) is not PdfNumber d ||
            ResolveObject(objects, matrixArray.Items[4]) is not PdfNumber e ||
            ResolveObject(objects, matrixArray.Items[5]) is not PdfNumber f) {
            return false;
        }

        matrix = new AppearancePlacement(a.Value, b.Value, c.Value, d.Value, e.Value, f.Value);
        return true;
    }

    private static bool TryInvertMatrix(AppearancePlacement matrix, out AppearancePlacement inverse) {
        double determinant = matrix.A * matrix.D - matrix.B * matrix.C;
        if (Math.Abs(determinant) < 0.0000001D) {
            inverse = new AppearancePlacement(1D, 0D, 0D, 1D, 0D, 0D);
            return false;
        }

        double a = matrix.D / determinant;
        double b = -matrix.B / determinant;
        double c = -matrix.C / determinant;
        double d = matrix.A / determinant;
        double e = (matrix.C * matrix.F - matrix.D * matrix.E) / determinant;
        double f = (matrix.B * matrix.E - matrix.A * matrix.F) / determinant;
        inverse = new AppearancePlacement(a, b, c, d, e, f);
        return true;
    }

    private static AppearancePlacement MultiplyMatrices(AppearancePlacement left, AppearancePlacement right) {
        return new AppearancePlacement(
            left.A * right.A + left.B * right.C,
            left.A * right.B + left.B * right.D,
            left.C * right.A + left.D * right.C,
            left.C * right.B + left.D * right.D,
            left.E * right.A + left.F * right.C + right.E,
            left.E * right.B + left.F * right.D + right.F);
    }

    private static bool TryReadBoxCoordinates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key, out double x, out double y, out double width, out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (!dictionary.Items.TryGetValue(key, out var boxObject) ||
            ResolveObject(objects, boxObject) is not PdfArray box ||
            box.Items.Count < 4 ||
            ResolveObject(objects, box.Items[0]) is not PdfNumber x1 ||
            ResolveObject(objects, box.Items[1]) is not PdfNumber y1 ||
            ResolveObject(objects, box.Items[2]) is not PdfNumber x2 ||
            ResolveObject(objects, box.Items[3]) is not PdfNumber y2) {
            return false;
        }

        x = Math.Min(x1.Value, x2.Value);
        y = Math.Min(y1.Value, y2.Value);
        width = Math.Abs(x2.Value - x1.Value);
        height = Math.Abs(y2.Value - y1.Value);
        return width > 0D && height > 0D;
    }

    private static bool TryGetNormalAppearanceReference(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, out PdfReference? reference) {
        reference = null;
        if (ResolveDictionary(objects, annotation.Items.TryGetValue("AP", out var appearanceObject) ? appearanceObject : null) is not PdfDictionary appearance ||
            !appearance.Items.TryGetValue("N", out var normalAppearanceObject)) {
            return false;
        }

        if (normalAppearanceObject is PdfReference normalAppearanceReference) {
            reference = normalAppearanceReference;
            return true;
        }

        if (ResolveDictionary(objects, normalAppearanceObject) is not PdfDictionary normalAppearanceStates) {
            return false;
        }

        string? selectedState = TryReadName(objects, annotation, "AS");
        if (selectedState != null &&
            selectedState.Length > 0 &&
            normalAppearanceStates.Items.TryGetValue(selectedState, out var selectedAppearance) &&
            selectedAppearance is PdfReference selectedReference) {
            reference = selectedReference;
            return true;
        }

        foreach (var state in normalAppearanceStates.Items.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            if (string.Equals(state.Key, "Off", StringComparison.Ordinal)) {
                continue;
            }

            if (state.Value is PdfReference stateReference) {
                reference = stateReference;
                return true;
            }
        }

        return false;
    }

    private static string? TryReadString(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfStringObj text
            ? text.Value
            : null;
    }

    private static double? TryReadNumber(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfNumber number
            ? number.Value
            : null;
    }

    private static PdfColor? TryReadColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value) ||
            ResolveObject(objects, value) is not PdfArray color ||
            color.Items.Count < 3 ||
            ResolveObject(objects, color.Items[0]) is not PdfNumber r ||
            ResolveObject(objects, color.Items[1]) is not PdfNumber g ||
            ResolveObject(objects, color.Items[2]) is not PdfNumber b) {
            return null;
        }

        return new PdfColor(ClampColor(r.Value), ClampColor(g.Value), ClampColor(b.Value));
    }

    private static double ClampColor(double value) {
        if (value < 0D) {
            return 0D;
        }

        return value > 1D ? 1D : value;
    }
}
