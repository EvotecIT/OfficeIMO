namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Builds a safe, immutable, bounded projection of the active raw object graph.</summary>
    public PdfRawDocumentView RawStructure(PdfRawStructureOptions? options = null) {
        DemandContentExtraction("raw object");
        PdfRawStructureOptions effective = options ?? new PdfRawStructureOptions();
        int take = Math.Min(_objects.Count, effective.MaxObjects);
        var projected = new List<PdfRawObjectView>(take);
        foreach (KeyValuePair<int, PdfIndirectObject> pair in _objects.OrderBy(static pair => pair.Key).Take(take)) {
            PdfIndirectObject indirect = pair.Value;
            projected.Add(new PdfRawObjectView(indirect.ObjectNumber, indirect.Generation, ProjectRawValue(indirect.Value, effective, depth: 0)));
        }

        PdfDictionary? catalog = FindCatalog();
        int? catalogObjectNumber = null;
        if (catalog != null) {
            foreach (KeyValuePair<int, PdfIndirectObject> pair in _objects) {
                if (ReferenceEquals(pair.Value.Value, catalog)) {
                    catalogObjectNumber = pair.Key;
                    break;
                }
            }
        }

        string trailer = _trailerRaw.Length <= effective.MaxTextLength
            ? _trailerRaw
            : _trailerRaw.Substring(0, effective.MaxTextLength);
        return new PdfRawDocumentView(
            projected.AsReadOnly(),
            _objects.Count,
            catalogObjectNumber,
            trailer,
            _objects.Count > take || _trailerRaw.Length > effective.MaxTextLength,
            Security.Revisions);
    }

    private static PdfRawValue ProjectRawValue(PdfObject value, PdfRawStructureOptions options, int depth) {
        if (depth >= options.MaxDepth && (value is PdfArray || value is PdfDictionary || value is PdfStream)) {
            return new PdfRawValue(PdfRawValueKind.Truncated, isTruncated: true);
        }

        switch (value) {
            case PdfNull:
                return new PdfRawValue(PdfRawValueKind.Null);
            case PdfNumber number:
                return new PdfRawValue(PdfRawValueKind.Number, number: number.Value);
            case PdfBoolean boolean:
                return new PdfRawValue(PdfRawValueKind.Boolean, boolean: boolean.Value);
            case PdfName name:
                return new PdfRawValue(PdfRawValueKind.Name, text: BoundText(name.Name, options.MaxTextLength));
            case PdfStringObj text:
                return new PdfRawValue(PdfRawValueKind.TextString, text: BoundText(text.Value, options.MaxTextLength), isTruncated: text.Value.Length > options.MaxTextLength);
            case PdfReference reference:
                return new PdfRawValue(PdfRawValueKind.Reference, referenceObjectNumber: reference.ObjectNumber, referenceGeneration: reference.Generation);
            case PdfArray array:
                return ProjectRawArray(array, options, depth);
            case PdfDictionary dictionary:
                return ProjectRawDictionary(dictionary, options, depth, PdfRawValueKind.Dictionary);
            case PdfStream stream:
                PdfRawValue dictionaryView = ProjectRawDictionary(stream.Dictionary, options, depth, PdfRawValueKind.Stream);
                return new PdfRawValue(
                    PdfRawValueKind.Stream,
                    entries: dictionaryView.Entries,
                    streamLength: stream.Data.Length,
                    streamDecodingFailed: stream.DecodingFailed,
                    isTruncated: dictionaryView.IsTruncated);
            default:
                return new PdfRawValue(PdfRawValueKind.Truncated, isTruncated: true);
        }
    }

    private static PdfRawValue ProjectRawArray(PdfArray array, PdfRawStructureOptions options, int depth) {
        int take = Math.Min(array.Items.Count, options.MaxCollectionItems);
        var items = new List<PdfRawValue>(take);
        for (int i = 0; i < take; i++) {
            items.Add(ProjectRawValue(array.Items[i], options, depth + 1));
        }

        return new PdfRawValue(PdfRawValueKind.Array, items: items.AsReadOnly(), isTruncated: array.Items.Count > take);
    }

    private static PdfRawValue ProjectRawDictionary(PdfDictionary dictionary, PdfRawStructureOptions options, int depth, PdfRawValueKind kind) {
        int take = Math.Min(dictionary.Items.Count, options.MaxCollectionItems);
        var entries = new Dictionary<string, PdfRawValue>(take, StringComparer.Ordinal);
        int index = 0;
        foreach (KeyValuePair<string, PdfObject> pair in dictionary.Items.OrderBy(static pair => pair.Key, StringComparer.Ordinal)) {
            if (index++ >= take) break;
            entries[BoundText(pair.Key, options.MaxTextLength)] = ProjectRawValue(pair.Value, options, depth + 1);
        }

        return new PdfRawValue(
            kind,
            entries: new System.Collections.ObjectModel.ReadOnlyDictionary<string, PdfRawValue>(entries),
            isTruncated: dictionary.Items.Count > take);
    }

    private static string BoundText(string text, int maximumLength) {
        return text.Length <= maximumLength ? text : text.Substring(0, maximumLength);
    }
}
