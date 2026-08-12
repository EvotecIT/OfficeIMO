namespace OfficeIMO.Pdf;

/// <summary>Shared fail-closed semantics for calibrated color-space declarations.</summary>
internal static class PdfCalibratedColorSpaceSemantics {
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
