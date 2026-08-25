using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Serializes the active catalog-rooted object graph into a normalized full-rewrite PDF.</summary>
internal static class PdfDocumentObjectGraphRewriter {
    internal static byte[] Rewrite(
        byte[] sourcePdf,
        PdfReadOptions? sourceReadOptions,
        PdfStandardEncryptionOptions? outputEncryption,
        Func<Dictionary<int, PdfIndirectObject>, PdfDocumentSecurityInfo, int?>? mutateObjectGraph = null,
        long? maximumOutputBytes = null) {
        Guard.NotNull(sourcePdf, nameof(sourcePdf));
        if (maximumOutputBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(sourcePdf, sourceReadOptions);
        var parsed = PdfSyntax.ParseObjects(sourcePdf, sourceReadOptions);
        Dictionary<int, PdfIndirectObject> objects = parsed.Map;
        byte[]? permanentFileId = outputEncryption == null
            ? PdfSyntax.ReadPermanentTrailerIdentifier(parsed.TrailerRaw)
            : null;
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
        bool requiresPdf20 = reachableObjectNumbers.Any(objectNumber =>
            objects[objectNumber].Value is PdfDictionary dictionary &&
            string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) &&
            string.Equals(dictionary.Get<PdfName>("Tabs")?.Name, "A", StringComparison.Ordinal));
        bool requiresPdf16 = reachableObjectNumbers.Any(objectNumber =>
                objects[objectNumber].Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Subtype")?.Name == "OpenType");
        bool requiresPdf15 = reachableObjectNumbers.Any(objectNumber =>
            objects[objectNumber].Value is PdfDictionary dictionary &&
            dictionary.Get<PdfNumber>("Ff") is PdfNumber flags &&
            ((int)flags.Value & (16777216 | 33554432 | 67108864)) != 0) ||
            reachableObjectNumbers.Any(objectNumber =>
                objects[objectNumber].Value is PdfDictionary dictionary &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) &&
                dictionary.Get<PdfName>("Tabs") is not null);
        PdfFileVersion minimumVersion = requiresPdf20 ? PdfFileVersion.Pdf20 : requiresPdf16 ? PdfFileVersion.Pdf16 : requiresPdf15 ? PdfFileVersion.Pdf15 : PdfFileVersion.Pdf14;
        if (fileVersion < minimumVersion && !CatalogDeclaresAtLeast(root.Value as PdfDictionary, objects, minimumVersion)) {
            fileVersion = PdfFileAssembler.RequireAtLeast(fileVersion, minimumVersion);
            if (root.Value is PdfDictionary catalog && catalog.Items.ContainsKey("Version")) {
                catalog.Items["Version"] = new PdfName(PdfFileAssembler.GetHeaderVersion(minimumVersion));
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
        if (maximumOutputBytes.HasValue) {
            return RewriteBounded(
                objects,
                reachableObjectNumbers,
                context,
                numberMap[rootObjectNumber],
                infoObjectNumber.HasValue ? numberMap[infoObjectNumber.Value] : 0,
                fileVersion,
                outputEncryption,
                permanentFileId,
                maximumOutputBytes.Value);
        }

        var serializedObjects = new List<byte[]>(reachableObjectNumbers.Count);
        long serializedObjectBytes = 0L;
        for (int i = 0; i < reachableObjectNumbers.Count; i++) {
            int sourceObjectNumber = reachableObjectNumbers[i];
            byte[] body = PdfPageExtractor.SerializeObject(objects[sourceObjectNumber].Value, context);
            byte[] serializedObject = PdfObjectBytes.WrapIndirectObject(i + 1, body);
            serializedObjectBytes = AddWithinOutputLimit(serializedObjectBytes, serializedObject.LongLength, maximumOutputBytes: null);
            serializedObjects.Add(serializedObject);
        }

        int rewrittenRootObjectNumber = numberMap[rootObjectNumber];
        int rewrittenInfoObjectNumber = infoObjectNumber.HasValue ? numberMap[infoObjectNumber.Value] : 0;
        return permanentFileId == null
            ? PdfFileAssembler.Assemble(
                serializedObjects,
                rewrittenRootObjectNumber,
                rewrittenInfoObjectNumber,
                fileVersion,
                outputEncryption)
            : PdfFileAssembler.AssemblePreservingPermanentId(
                serializedObjects,
                rewrittenRootObjectNumber,
                rewrittenInfoObjectNumber,
                fileVersion,
                outputEncryption,
                permanentFileId);
    }

    private static byte[] RewriteBounded(
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<int> reachableObjectNumbers,
        PdfPageExtractor.SerializationContext context,
        int rewrittenRootObjectNumber,
        int rewrittenInfoObjectNumber,
        PdfFileVersion fileVersion,
        PdfStandardEncryptionOptions? outputEncryption,
        byte[]? permanentFileId,
        long maximumOutputBytes) {
        using var serializedObjects = new PdfObjectStore(memoryLimitBytes: 0L);
        long serializedObjectBytes = 0L;
        byte[] suffix = PdfEncoding.Latin1GetBytes("endobj\n");
        for (int i = 0; i < reachableObjectNumbers.Count; i++) {
            int rewrittenObjectNumber = i + 1;
            int sourceObjectNumber = reachableObjectNumbers[i];
            long remaining = maximumOutputBytes - serializedObjectBytes;
            PdfPageExtractor.EnsureSerializedIndirectObjectWithinLimit(
                objects[sourceObjectNumber].Value,
                context,
                rewrittenObjectNumber,
                remaining);
            byte[] body = PdfPageExtractor.SerializeObject(objects[sourceObjectNumber].Value, context);
            byte[] prefix = PdfEncoding.Latin1GetBytes(
                rewrittenObjectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
            long objectBytes = AddWithinOutputLimit(prefix.LongLength, body.LongLength, maximumOutputBytes);
            objectBytes = AddWithinOutputLimit(objectBytes, suffix.LongLength, maximumOutputBytes);
            serializedObjectBytes = AddWithinOutputLimit(serializedObjectBytes, objectBytes, maximumOutputBytes);
            serializedObjects.AddSegments(prefix, body, suffix);
        }

        using FileStream output = PdfTemporaryFile.Create(".rewrite", FileOptions.RandomAccess, out _);
        using var boundedOutput = new PdfBoundedWriteStream(
            output,
            maximumOutputBytes,
            "The rewritten PDF exceeds the configured expanded container limit.");
        if (permanentFileId == null) {
            PdfFileAssembler.Assemble(
                boundedOutput,
                serializedObjects,
                rewrittenRootObjectNumber,
                rewrittenInfoObjectNumber,
                fileVersion,
                outputEncryption,
                objectMemoryLimitBytes: 0L);
        } else {
            PdfFileAssembler.AssemblePreservingPermanentId(
                boundedOutput,
                serializedObjects,
                rewrittenRootObjectNumber,
                rewrittenInfoObjectNumber,
                fileVersion,
                outputEncryption,
                permanentFileId,
                objectMemoryLimitBytes: 0L);
        }
        boundedOutput.Flush();
        return ReadBoundedOutput(output, maximumOutputBytes);
    }

    private static byte[] ReadBoundedOutput(FileStream output, long maximumOutputBytes) {
        ThrowIfOutputLimitExceeded(output.Length, maximumOutputBytes);
        if (output.Length > int.MaxValue) {
            throw new InvalidDataException("The rewritten PDF exceeds the supported in-memory result size.");
        }
        var bytes = new byte[(int)output.Length];
        output.Position = 0L;
        int read = 0;
        while (read < bytes.Length) {
            int count = output.Read(bytes, read, bytes.Length - read);
            if (count == 0) throw new EndOfStreamException("The temporary rewritten PDF ended unexpectedly.");
            read += count;
        }
        return bytes;
    }

    private static long AddWithinOutputLimit(long current, long added, long? maximumOutputBytes) {
        long total = current > long.MaxValue - added ? long.MaxValue : current + added;
        ThrowIfOutputLimitExceeded(total, maximumOutputBytes);
        return total;
    }

    private static void ThrowIfOutputLimitExceeded(long observedBytes, long? maximumOutputBytes) {
        if (maximumOutputBytes.HasValue && observedBytes > maximumOutputBytes.Value) {
            throw new InvalidDataException("The rewritten PDF exceeds the configured expanded container limit.");
        }
    }

    private static bool CatalogDeclaresAtLeast(
        PdfDictionary? catalog,
        Dictionary<int, PdfIndirectObject> objects,
        PdfFileVersion minimumVersion) {
        string? version = catalog != null &&
            catalog.Items.TryGetValue("Version", out PdfObject? versionObject) &&
            TryResolveReferenceChain(objects, versionObject, out PdfObject? resolvedVersion) &&
            resolvedVersion is PdfName versionName
                ? versionName.Name
                : null;
        return version != null && PdfFileAssembler.ParseHeaderVersionOrDefault(version) >= minimumVersion;
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
