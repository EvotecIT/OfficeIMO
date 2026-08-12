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
        var serializedObjects = new List<byte[]>(reachableObjectNumbers.Count);
        long serializedObjectBytes = 0L;
        for (int i = 0; i < reachableObjectNumbers.Count; i++) {
            int sourceObjectNumber = reachableObjectNumbers[i];
            if (maximumOutputBytes.HasValue) {
                long remaining = maximumOutputBytes.Value - serializedObjectBytes;
                PdfPageExtractor.EnsureSerializedIndirectObjectWithinLimit(
                    objects[sourceObjectNumber].Value, context, i + 1, remaining);
            }
            byte[] body = PdfPageExtractor.SerializeObject(objects[sourceObjectNumber].Value, context);
            ThrowIfOutputLimitExceeded(body.LongLength, maximumOutputBytes);
            byte[] serializedObject = PdfObjectBytes.WrapIndirectObject(i + 1, body);
            serializedObjectBytes = AddWithinOutputLimit(serializedObjectBytes, serializedObject.LongLength, maximumOutputBytes);
            serializedObjects.Add(serializedObject);
        }

        int rewrittenRootObjectNumber = numberMap[rootObjectNumber];
        int rewrittenInfoObjectNumber = infoObjectNumber.HasValue ? numberMap[infoObjectNumber.Value] : 0;
        if (!maximumOutputBytes.HasValue) {
            return PdfFileAssembler.Assemble(
                serializedObjects,
                rewrittenRootObjectNumber,
                rewrittenInfoObjectNumber,
                fileVersion,
                outputEncryption);
        }
        using var output = new MemoryStream();
        using var boundedOutput = new PdfBoundedWriteStream(output, maximumOutputBytes);
        PdfFileAssembler.Assemble(
            boundedOutput,
            serializedObjects,
            rewrittenRootObjectNumber,
            rewrittenInfoObjectNumber,
            fileVersion,
            outputEncryption);
        return output.ToArray();
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

    private sealed class PdfBoundedWriteStream : Stream {
        private readonly Stream _inner;
        private readonly long? _maximumBytes;

        internal PdfBoundedWriteStream(Stream inner, long? maximumBytes) {
            _inner = inner;
            _maximumBytes = maximumBytes;
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) {
            AddWithinOutputLimit(_inner.Position, count, _maximumBytes);
            _inner.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value) {
            AddWithinOutputLimit(_inner.Position, 1L, _maximumBytes);
            _inner.WriteByte(value);
        }

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Flush();
            base.Dispose(disposing);
        }
    }
}
