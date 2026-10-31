namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    /// <summary>Writes the pruned live graph, preserving null semantics for undefined references.</summary>
    private static byte[] RewriteAllObjects(
        Dictionary<int, PdfIndirectObject> objects,
        int catalogObjectNumber,
        PdfMetadata metadata,
        byte[] sourcePdf,
        PdfMutationPlan? mutationPlan = null,
        PdfRedactionSearchCriteria[]? searchCriteria = null,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int[] sourceIds = objects.Keys.OrderBy(id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            if ((i & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            numberMap[sourceIds[i]] = i + 1;
        }

        if (mutationPlan is { Operation: PdfMutationOperation.Redact }) {
            var collector = new PdfPageExtractor.ObjectCollector(objects, undefinedReferencesAreNull: true, cancellationToken: cancellationToken);
            try {
                collector.CollectObjectGraph(new PdfReference(catalogObjectNumber, objects[catalogObjectNumber].Generation));
            } catch (InvalidOperationException) {
                throw PdfMutationBlockedException.FailedProof(mutationPlan, PdfMutationProof.RewritePreservation);
            }
            // Source undefined references were normalized before allocation. New dangling targets are unproven.
            if (collector.UndefinedObjectNumbers.Count > 0)
                throw PdfMutationBlockedException.FailedProof(mutationPlan, PdfMutationProof.RewritePreservation);
        }
        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(), objects, cancellationToken: cancellationToken);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            cancellationToken.ThrowIfCancellationRequested();
            rewritten.Add(PdfPageExtractor.SerializeIndirectObject(numberMap[sourceId], objects[sourceId].Value, context));
        }

        cancellationToken.ThrowIfCancellationRequested();
        int infoId = rewritten.Count + 1;
        PdfDictionary info = PdfInfoDictionaryBuilder.BuildDictionary(metadata);
        // Generated producer metadata must not reintroduce a short searched value either.
        foreach (string key in info.Items.Keys.ToArray()) {
            if (info.Items[key] is PdfStringObj text && searchCriteria?.Any(criterion => criterion.Matches(text.Value)) == true) info.Items.Remove(key);
        }
        rewritten.Add(PdfPageExtractor.SerializeIndirectObject(infoId, info, context));

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion, cancellationToken);
    }
}
