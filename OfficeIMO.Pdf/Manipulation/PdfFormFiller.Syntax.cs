namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
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

    private static string? TryReadText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfStringObj text &&
            !string.IsNullOrEmpty(text.Value)
            ? text.Value
            : null;
    }

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static string? CombineFieldName(string? parentName, string? partialName) {
        if (string.IsNullOrEmpty(parentName)) {
            return string.IsNullOrEmpty(partialName) ? null : partialName;
        }

        if (string.IsNullOrEmpty(partialName)) {
            return parentName;
        }

        return parentName + "." + partialName;
    }
}
