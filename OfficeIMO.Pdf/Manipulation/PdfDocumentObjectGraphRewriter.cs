namespace OfficeIMO.Pdf;

/// <summary>Serializes the active catalog-rooted object graph into a normalized full-rewrite PDF.</summary>
internal static class PdfDocumentObjectGraphRewriter {
    internal static byte[] Rewrite(
        byte[] sourcePdf,
        PdfReadOptions? sourceReadOptions,
        PdfStandardEncryptionOptions? outputEncryption,
        Func<Dictionary<int, PdfIndirectObject>, PdfDocumentSecurityInfo, int?>? mutateObjectGraph = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(sourcePdf, sourceReadOptions);
        var parsed = PdfSyntax.ParseObjects(sourcePdf, sourceReadOptions);
        Dictionary<int, PdfIndirectObject> objects = parsed.Map;
        int rootObjectNumber = RequireRootObjectNumber(security, objects);
        int? infoObjectNumber = mutateObjectGraph is null
            ? FindInfoObjectNumber(security, objects)
            : mutateObjectGraph(objects, security);
        rootObjectNumber = RequireRootObjectNumber(security, objects);

        var collector = new PdfPageExtractor.ObjectCollector(objects);
        PdfIndirectObject root = objects[rootObjectNumber];
        collector.CollectObjectGraph(new PdfReference(root.ObjectNumber, root.Generation));
        if (infoObjectNumber.HasValue) {
            PdfIndirectObject info = objects[infoObjectNumber.Value];
            collector.CollectObjectGraph(new PdfReference(info.ObjectNumber, info.Generation));
        }

        IReadOnlyList<int> reachableObjectNumbers = collector.ObjectIds;
        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        if (reachableObjectNumbers.Any(objectNumber =>
                objects[objectNumber].Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Subtype")?.Name == "OpenType") &&
            !CatalogDeclaresAtLeastPdf16(root.Value as PdfDictionary, objects)) {
            fileVersion = PdfFileAssembler.RequireAtLeast(fileVersion, PdfFileVersion.Pdf16);
            if (root.Value is PdfDictionary catalog && catalog.Items.ContainsKey("Version")) {
                catalog.Items["Version"] = new PdfName("1.6");
                collector = new PdfPageExtractor.ObjectCollector(objects);
                collector.CollectObjectGraph(new PdfReference(root.ObjectNumber, root.Generation));
                if (infoObjectNumber.HasValue) {
                    PdfIndirectObject info = objects[infoObjectNumber.Value];
                    collector.CollectObjectGraph(new PdfReference(info.ObjectNumber, info.Generation));
                }
                reachableObjectNumbers = collector.ObjectIds;
            }
        }
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

        int rewrittenRootObjectNumber = numberMap[rootObjectNumber];
        int rewrittenInfoObjectNumber = infoObjectNumber.HasValue ? numberMap[infoObjectNumber.Value] : 0;
        return PdfFileAssembler.Assemble(
            serializedObjects,
            rewrittenRootObjectNumber,
            rewrittenInfoObjectNumber,
            fileVersion,
            outputEncryption);
    }

    private static bool CatalogDeclaresAtLeastPdf16(
        PdfDictionary? catalog,
        Dictionary<int, PdfIndirectObject> objects) {
        string? version = catalog != null &&
            catalog.Items.TryGetValue("Version", out PdfObject? versionObject) &&
            TryResolveReferenceChain(objects, versionObject, out PdfObject? resolvedVersion) &&
            resolvedVersion is PdfName versionName
                ? versionName.Name
                : null;
        return version == "1.6" || version == "1.7" || version == "2.0";
    }

    private static bool TryResolveReferenceChain(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        out PdfObject? resolved) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect) ||
                indirect.Generation != reference.Generation) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return true;
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
