namespace OfficeIMO.Pdf;

/// <summary>Serializes the active catalog-rooted object graph into a normalized full-rewrite PDF.</summary>
internal static class PdfDocumentObjectGraphRewriter {
    internal static byte[] Rewrite(
        byte[] sourcePdf,
        PdfReadOptions? sourceReadOptions,
        PdfStandardEncryptionOptions? outputEncryption) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(sourcePdf, sourceReadOptions);
        var parsed = PdfSyntax.ParseObjects(sourcePdf, sourceReadOptions);
        Dictionary<int, PdfIndirectObject> objects = parsed.Map;
        int rootObjectNumber = RequireRootObjectNumber(security, objects);
        int? infoObjectNumber = FindInfoObjectNumber(security, objects);

        var collector = new PdfPageExtractor.ObjectCollector(objects);
        PdfIndirectObject root = objects[rootObjectNumber];
        collector.CollectObjectGraph(new PdfReference(root.ObjectNumber, root.Generation));
        if (infoObjectNumber.HasValue) {
            PdfIndirectObject info = objects[infoObjectNumber.Value];
            collector.CollectObjectGraph(new PdfReference(info.ObjectNumber, info.Generation));
        }

        IReadOnlyList<int> reachableObjectNumbers = collector.ObjectIds;
        var numberMap = new Dictionary<int, int>(reachableObjectNumbers.Count);
        for (int i = 0; i < reachableObjectNumbers.Count; i++) {
            numberMap[reachableObjectNumbers[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(
            numberMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            objects,
            preserveRawStringBytes: true);
        var serializedObjects = new List<byte[]>(reachableObjectNumbers.Count);
        for (int i = 0; i < reachableObjectNumbers.Count; i++) {
            int sourceObjectNumber = reachableObjectNumbers[i];
            byte[] body = PdfPageExtractor.SerializeObject(objects[sourceObjectNumber].Value, context);
            serializedObjects.Add(PdfObjectBytes.WrapIndirectObject(i + 1, body));
        }

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        int rewrittenRootObjectNumber = numberMap[rootObjectNumber];
        int rewrittenInfoObjectNumber = infoObjectNumber.HasValue ? numberMap[infoObjectNumber.Value] : 0;
        return PdfFileAssembler.Assemble(
            serializedObjects,
            rewrittenRootObjectNumber,
            rewrittenInfoObjectNumber,
            fileVersion,
            outputEncryption);
    }

    private static int RequireRootObjectNumber(
        PdfDocumentSecurityInfo security,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) ||
            root.Value is not PdfDictionary) {
            throw new InvalidOperationException("The active PDF trailer does not reference a readable catalog object.");
        }

        return security.RootObjectNumber.Value;
    }

    private static int? FindInfoObjectNumber(
        PdfDocumentSecurityInfo security,
        Dictionary<int, PdfIndirectObject> objects) {
        return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
            ? security.InfoObjectNumber
            : null;
    }
}
