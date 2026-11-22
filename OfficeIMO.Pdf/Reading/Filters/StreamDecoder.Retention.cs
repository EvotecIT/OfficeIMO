namespace OfficeIMO.Pdf.Filters;

internal static partial class StreamDecoder {
    // Escaped metadata needs the decoder's inputs to remain available. Actual
    // decoding still runs under its normal limits when a profile property is read.
    internal static void RetainEncodedInput(PdfStream stream, Dictionary<int, PdfIndirectObject> objects,
        PdfIccProfileRetentionBudget budget, CancellationToken cancellationToken) {
        var pending = new Queue<(PdfStream Stream, int Depth)>();
        var seen = new HashSet<PdfStream> { stream };
        pending.Enqueue((stream, 0));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            (PdfStream current, int depth) = pending.Dequeue();
            budget.Charge(current, PdfIccProfileCacheRepresentation.EncodedBytes, current.DataLength);
            _ = current.GetData(cancellationToken);
            PdfDictionary dictionary = current.Dictionary;
            if (!dictionary.Items.TryGetValue("Filter", out PdfObject? filter) ||
                !TryGetFilterNames(filter, objects, out List<string> names, cancellationToken) ||
                !HasValidDecodeParmsDeclaration(dictionary, names.Count, objects, cancellationToken) ||
                !HasValidDecodeParmsForFilters(dictionary, names, objects, cancellationToken)) {
                continue;
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
                    ResolveObject(value, objects, cancellationToken) is PdfStream globals && seen.Add(globals)) {
                    pending.Enqueue((globals, depth + 1));
                }
            }
        }
    }
}
