namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private PdfPortfolioInfo? ExtractPortfolio() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog == null || !catalog.Items.TryGetValue("Collection", out PdfObject? collectionObject) ||
            ResolveObject(collectionObject) is not PdfDictionary collection) {
            return null;
        }

        string? view = ReadName(collection, "View");
        string? initialDocument = ReadString(collection, "D");
        var fields = new List<PdfPortfolioFieldInfo>();
        if (collection.Items.TryGetValue("Schema", out PdfObject? schemaObject) && ResolveObject(schemaObject) is PdfDictionary schema) {
            foreach (KeyValuePair<string, PdfObject> entry in schema.Items) {
                if (ResolveObject(entry.Value) is not PdfDictionary field) continue;
                fields.Add(new PdfPortfolioFieldInfo(
                    entry.Key,
                    ReadString(field, "N"),
                    ReadName(field, "Subtype"),
                    ReadInteger(field, "O"),
                    ReadBoolean(field, "V"),
                    ReadBoolean(field, "E")));
            }
            fields.Sort((left, right) => Nullable.Compare(left.Order, right.Order));
        }

        string? sortField = null;
        bool? sortAscending = null;
        if (collection.Items.TryGetValue("Sort", out PdfObject? sortObject) && ResolveObject(sortObject) is PdfDictionary sort) {
            sortField = ReadName(sort, "S");
            sortAscending = ReadBoolean(sort, "A");
        }

        return new PdfPortfolioInfo(view, initialDocument, fields.AsReadOnly(), sortField, sortAscending);
    }

    private string? ReadName(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value) is PdfName name ? name.Name : null;

    private string? ReadString(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value) is PdfStringObj text ? text.Value : null;

    private int? ReadInteger(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value) is PdfNumber number &&
        number.Value >= int.MinValue && number.Value <= int.MaxValue && number.Value == Math.Truncate(number.Value)
            ? (int)number.Value
            : null;

    private bool? ReadBoolean(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && ResolveObject(value) is PdfBoolean boolean ? boolean.Value : null;
}
