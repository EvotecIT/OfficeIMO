namespace OfficeIMO.Pdf;

/// <summary>Shared fail-closed semantics for calibrated color-space declarations.</summary>
internal static class PdfCalibratedColorSpaceSemantics {
    internal static bool IsStructurallyValid(
        string family,
        PdfDictionary calibration,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (maximumObjectDepth < 0 ||
            !TryReadNumberArray(calibration, "WhitePoint", 3, objects, maximumObjectDepth, out double[] whitePoint) ||
            !IsValidWhitePoint(whitePoint) ||
            !HasValidBlackPoint(calibration, objects, maximumObjectDepth)) return false;

        switch (family) {
            case "CalGray":
                return TryReadOptionalPositiveNumber(calibration, "Gamma", objects, maximumObjectDepth);
            case "CalRGB":
                return TryReadOptionalNumberArray(
                        calibration, "Gamma", 3, objects, maximumObjectDepth,
                        static values => values.All(static value => value > 0D)) &&
                    TryReadOptionalNumberArray(
                        calibration, "Matrix", 9, objects, maximumObjectDepth,
                        static _ => true);
            case "Lab":
                return TryReadOptionalNumberArray(
                    calibration, "Range", 4, objects, maximumObjectDepth, IsSupportedLabRange);
            default:
                return false;
        }
    }

    internal static bool IsValidWhitePoint(IReadOnlyList<double> values) =>
        values.Count == 3 &&
        values[0] > 0D && values[1] == 1D && values[2] > 0D &&
        !double.IsNaN(values[0]) && !double.IsInfinity(values[0]) &&
        !double.IsNaN(values[2]) && !double.IsInfinity(values[2]);

    internal static bool IsSupportedLabRange(IReadOnlyList<double> values) =>
        values.Count == 4 &&
        values[0] >= -128D && values[1] <= 127D &&
        values[2] >= -128D && values[3] <= 127D &&
        values[0] < values[1] && values[2] < values[3] &&
        values.All(static value => !double.IsNaN(value) && !double.IsInfinity(value));

    internal static bool HasSupportedBlackPoint(
        PdfDictionary calibration,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!calibration.Items.TryGetValue("BlackPoint", out PdfObject? value)) return true;
        if (!TryResolve(value, objects, out PdfObject? resolved)) return false;
        if (resolved is PdfNull) return true;
        if (resolved is not PdfArray { Items.Count: 3 } array) return false;

        for (int index = 0; index < array.Items.Count; index++) {
            if (!TryResolve(array.Items[index], objects, out PdfObject? component) ||
                component is not PdfNumber number ||
                double.IsNaN(number.Value) || double.IsInfinity(number.Value) ||
                number.Value != 0D) return false;
        }
        return true;
    }

    private static bool HasValidBlackPoint(
        PdfDictionary calibration,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (!calibration.Items.TryGetValue("BlackPoint", out PdfObject? value)) return true;
        if (!TryResolve(value, objects, maximumObjectDepth, out PdfObject? resolved)) return false;
        if (resolved is PdfNull) return true;
        if (resolved is not PdfArray { Items.Count: 3 } array) return false;
        for (int index = 0; index < array.Items.Count; index++) {
            if (!TryResolve(array.Items[index], objects, maximumObjectDepth, out PdfObject? component) ||
                component is not PdfNumber number ||
                !IsFinite(number.Value) ||
                number.Value < 0D) return false;
        }
        return true;
    }

    private static bool TryReadOptionalPositiveNumber(
        PdfDictionary calibration,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (!calibration.Items.TryGetValue(key, out PdfObject? value)) return true;
        if (!TryResolve(value, objects, maximumObjectDepth, out PdfObject? resolved)) return false;
        if (resolved is PdfNull) return true;
        return resolved is PdfNumber number && IsFinite(number.Value) && number.Value > 0D;
    }

    private static bool TryReadOptionalNumberArray(
        PdfDictionary calibration,
        string key,
        int count,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        Func<IReadOnlyList<double>, bool> validator) {
        if (!calibration.Items.TryGetValue(key, out PdfObject? value)) return true;
        if (!TryResolve(value, objects, maximumObjectDepth, out PdfObject? resolved)) return false;
        if (resolved is PdfNull) return true;
        return TryReadNumberArray(resolved, count, objects, maximumObjectDepth, out double[] values) &&
            validator(values);
    }

    private static bool TryReadNumberArray(
        PdfDictionary dictionary,
        string key,
        int count,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out double[] values) {
        values = Array.Empty<double>();
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            TryReadNumberArray(value, count, objects, maximumObjectDepth, out values);
    }

    private static bool TryReadNumberArray(
        PdfObject? value,
        int count,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out double[] values) {
        values = Array.Empty<double>();
        if (!TryResolve(value, objects, maximumObjectDepth, out PdfObject? resolved) ||
            resolved is not PdfArray array || array.Items.Count != count) return false;
        values = new double[count];
        for (int index = 0; index < count; index++) {
            if (!TryResolve(array.Items[index], objects, maximumObjectDepth, out PdfObject? component) ||
                component is not PdfNumber number || !IsFinite(number.Value)) return false;
            values[index] = number.Value;
        }
        return true;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool TryResolve(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject? resolved) {
        return TryResolve(value, objects, int.MaxValue, out resolved);
    }

    private static bool TryResolve(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out PdfObject? resolved) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        int depth = 0;
        while (resolved is PdfReference reference) {
            if (depth++ >= maximumObjectDepth ||
                !visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return resolved != null;
    }
}
