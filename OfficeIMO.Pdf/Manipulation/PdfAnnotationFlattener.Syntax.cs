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

    private static Matrix2D ReadAppearancePlacement(Dictionary<int, PdfIndirectObject> objects, PdfReference appearanceReference, double x, double y, double width, double height) {
        if (!PdfObjectLookup.TryGet(objects, appearanceReference, out var appearanceObject) ||
            appearanceObject.Value is not PdfStream appearanceStream)
            throw new InvalidOperationException("The annotation appearance stream is missing.");
        return PdfAppearancePlacement.Read(appearanceStream.Dictionary, value => ResolveObject(objects, value),
            x, y, width, height, out _);
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
