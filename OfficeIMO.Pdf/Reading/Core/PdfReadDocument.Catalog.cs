namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private PdfDocumentOpenAction? ExtractOpenAction() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("OpenAction", out var openActionObject)) {
            return null;
        }

        PdfObject? resolved = ResolveObject(openActionObject);
        if (resolved is PdfArray &&
            TryReadDestination(resolved, out int? pageNumber, out double? destinationTop, out PdfOpenActionDestinationMode? destinationMode, out double? destinationLeft, out double? destinationBottom, out double? destinationRight, out double? destinationZoom)) {
            return new PdfDocumentOpenAction("Destination", pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight, destinationZoom);
        }

        if (resolved is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var destination) &&
            TryReadDestination(destination, out pageNumber, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight, out destinationZoom)) {
            return new PdfDocumentOpenAction("GoTo", pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight, destinationZoom);
        }

        return null;
    }

    private PdfViewerPreferences? ExtractViewerPreferences() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("ViewerPreferences", out var viewerPreferencesObject) ||
            ResolveObject(viewerPreferencesObject) is not PdfDictionary dictionary) {
            return null;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in dictionary.Items) {
            if (!TryFormatSimpleValue(entry.Value, out string? value)) {
                return null;
            }

            values[entry.Key] = value!;
        }

        return values.Count == 0 ? null : new PdfViewerPreferences(values);
    }

    private PdfDictionary? FindCatalog() {
        return PdfSyntax.FindCatalog(_objects, _trailerRaw);
    }

    private string? ExtractCatalogName(string key) {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfName name ||
            string.IsNullOrEmpty(name.Name)) {
            return null;
        }

        return name.Name;
    }

    private string? ExtractCatalogString(string key) {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfStringObj text ||
            string.IsNullOrEmpty(text.Value)) {
            return null;
        }

        return text.Value;
    }
}
