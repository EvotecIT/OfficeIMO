namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static PdfDictionary? ResolveFontDescriptor(
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects,
        out string? programFontSubtype) {
        PdfDictionary fontWithDescriptor = font;
        programFontSubtype = font.Get<PdfName>("Subtype")?.Name;
        if (string.Equals(programFontSubtype, "Type0", StringComparison.Ordinal) &&
            font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject)) {
            PdfArray? descendants = ResolveArray(descendantsObject, objects);
            PdfDictionary? descendant = descendants is { Items.Count: > 0 }
                ? ResolveDict(descendants.Items[0], objects)
                : null;
            if (descendant != null) {
                fontWithDescriptor = descendant;
                programFontSubtype = descendant.Get<PdfName>("Subtype")?.Name;
            }
        }

        return fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject)
            ? ResolveDict(descriptorObject, objects)
            : null;
    }

    private static int? TryReadFontDescriptorInteger(
        PdfDictionary? descriptor,
        Dictionary<int, PdfIndirectObject> objects,
        string key,
        int minimum,
        int maximum) {
        if (descriptor == null ||
            !descriptor.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(value, objects) is not PdfNumber number ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value) ||
            number.Value != Math.Truncate(number.Value) ||
            number.Value < minimum ||
            number.Value > maximum) return null;
        return (int)number.Value;
    }
}
