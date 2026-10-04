namespace OfficeIMO.IWork.Internal;

/// <summary>Reads bounded native style inheritance for text and table projections.</summary>
internal static class IWorkStyleReader {
    internal static IReadOnlyList<(IWorkArchiveRecord Record, IWorkWireMessage Message)> ReadChain(IWorkObjectIndex index,
        ulong identifier, int maximumDepth, Func<uint, bool> allowedType,
        bool tolerateStyleDepth, IWorkSourceReferenceIssueCollector references, ref bool complete,
        int superArchiveDepth = 1) {
        var chain = new List<(IWorkArchiveRecord Record, IWorkWireMessage Message)>();
        var seen = new HashSet<ulong>();
        ulong current = identifier;
        while (true) {
            if (chain.Count >= maximumDepth) {
                if (tolerateStyleDepth) {
                    complete = false;
                    break;
                }
                throw new InvalidDataException(
                    $"iWork text style inheritance exceeds the configured depth of {maximumDepth}.");
            }
            if (!seen.Add(current)) {
                complete = false;
                break;
            }
            IWorkArchiveRecord? record = index.Find(current);
            if (record == null || !allowedType(record.MessageType)) {
                complete = false;
                break;
            }
            IWorkWireMessage message;
            try {
                message = index.Message(record);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(record, "$", null);
                complete = false;
                break;
            }
            chain.Add((record, message));
            IWorkWireMessage? super = message;
            string superPath = string.Empty;
            for (int level = 0; level < superArchiveDepth && super != null; level++) {
                superPath = superPath.Length == 0 ? "1" : superPath + "/1";
                IWorkWireMessage? nested = IWorkObjectIndex.TryGetMessage(super, 1, out bool malformedSuper);
                if (malformedSuper || super.HasField(1) && nested == null) {
                    references.Declarations.Record(record, superPath, super.FieldCount(1));
                    complete = false;
                    super = null;
                    break;
                }
                super = nested;
            }
            if (super == null) break;
            IWorkArchiveRecord? parent = references.ReadOne(record, super, 3, superPath + "/3", allowedType);
            if (super.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                || super.HasField(3) && (parent == null || !allowedType(parent.MessageType))) {
                complete = false;
                break;
            }
            if (parent == null) break;
            current = parent.Identifier;
        }
        return chain;
    }

}
