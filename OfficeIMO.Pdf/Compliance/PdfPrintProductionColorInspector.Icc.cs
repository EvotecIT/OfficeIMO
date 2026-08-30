namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static bool HasValidIccBasedOptions(
        PdfDictionary profile,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        if (profile.Items.TryGetValue("Alternate", out PdfObject? alternateObject) &&
            ResolveObject(objects, alternateObject, 0, maximumObjectDepth) is not PdfNull) {
            ColorSpaceUsage alternate = ClassifyColorSpace(
                alternateObject,
                objects,
                maximumObjectDepth,
                maximumDecodedStreamBytes);
            if (!alternate.IsKnown || alternate.UsesPattern || alternate.ComponentCount != componentCount) return false;
        }

        if (!profile.Items.TryGetValue("Range", out PdfObject? rangeObject) ||
            ResolveObject(objects, rangeObject, 0, maximumObjectDepth) is PdfNull) return true;
        if (!HasExactFiniteNumberArray(
                profile,
                "Range",
                checked(componentCount * 2),
                objects,
                maximumObjectDepth,
                out double[] ranges)) return false;
        for (int component = 0; component < componentCount; component++) {
            if (ranges[component * 2] > ranges[component * 2 + 1]) return false;
        }
        return true;
    }
}
