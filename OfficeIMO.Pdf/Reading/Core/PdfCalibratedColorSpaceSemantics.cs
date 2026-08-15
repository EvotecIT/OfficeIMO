namespace OfficeIMO.Pdf;

/// <summary>Shared fail-closed semantics for calibrated color-space declarations.</summary>
internal static class PdfCalibratedColorSpaceSemantics {
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

    private static bool TryResolve(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfObject? resolved) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return resolved != null;
    }
}
