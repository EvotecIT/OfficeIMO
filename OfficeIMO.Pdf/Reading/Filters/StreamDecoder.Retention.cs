namespace OfficeIMO.Pdf.Filters;

internal static partial class StreamDecoder {
    // Metadata can outlive a borrowed source. Resolve only the decoder's inputs
    // while the source is still valid. Actual decoding keeps its normal budget.
    internal static void RetainEncodedInput(PdfStream stream, Dictionary<int, PdfIndirectObject> objects,
        CancellationToken cancellationToken) =>
        RetainEncodedInput(stream, objects, visited: null, depth: 0, cancellationToken);

    private static void RetainEncodedInput(PdfStream stream, Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfStream>? visited, int depth, CancellationToken cancellationToken) {
        _ = stream.GetData(cancellationToken);
        PdfDictionary dictionary = stream.Dictionary;
        if (!dictionary.Items.TryGetValue("Filter", out PdfObject? filter) ||
            !TryGetFilterNames(filter, objects, out List<string> names, cancellationToken) ||
            !HasValidDecodeParmsDeclaration(dictionary, names.Count, objects, cancellationToken) ||
            !HasValidDecodeParmsForFilters(dictionary, names, objects, cancellationToken)) {
            return;
        }

        for (int index = 0; index < names.Count; index++) {
            DecodeFilterKind kind = GetFilterKind(names[index]);
            if (kind is DecodeFilterKind.Fax or DecodeFilterKind.Jbig2) {
                _ = ResolveObject(dictionary.Get<PdfReference>("Width"), objects, cancellationToken);
                _ = ResolveObject(dictionary.Get<PdfReference>("Height"), objects, cancellationToken);
            }
            if (kind == DecodeFilterKind.Jbig2 && depth < MaximumJbig2GlobalsDepth &&
                GetDecodeParms(dictionary, index, objects, cancellationToken) is { } parameters &&
                parameters.Items.TryGetValue("JBIG2Globals", out PdfObject? value) &&
                ResolveObject(value, objects, cancellationToken) is PdfStream globals) {
                visited ??= new HashSet<PdfStream> { stream };
                if (visited.Add(globals)) RetainEncodedInput(globals, objects, visited, depth + 1, cancellationToken);
            }
        }
    }
}
