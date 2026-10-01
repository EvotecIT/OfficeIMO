namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static IReadOnlyDictionary<int, IWorkInlineObject> ReadInlineObjects(IWorkObjectIndex index,
        IWorkWireMessage message, string text, IWorkArchiveRecord storage, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, out bool complete) {
        complete = true;
        var boundaries = ReadObjectTable(message, 9, text.Length, storage, budget, references, ref complete);
        var result = new Dictionary<int, IWorkInlineObject>();
        var duplicateOffsets = new HashSet<int>(boundaries.GroupBy(boundary => boundary.Index)
            .Where(group => group.Count() > 1).Select(group => group.Key));
        foreach (AttributeBoundary boundary in boundaries) {
            if (!boundary.HasObject) continue;
            string path = FormattableString.Invariant($"9/1[{boundary.Position}]");
            if (duplicateOffsets.Contains(boundary.Index) || boundary.Index >= text.Length || text[boundary.Index] != '\ufffc') {
                complete = false;
                references.Declarations.Record(storage, path, 1, IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                continue;
            }
            IWorkArchiveRecord? attachment = boundary.Identifier.HasValue ? index.Find(boundary.Identifier.Value) : null;
            if (attachment?.MessageType != 2003) {
                complete = false;
                if (attachment != null) references.Declarations.Record(attachment, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
                continue;
            }
            IWorkWireMessage payload;
            try { payload = index.Message(attachment); }
            catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                complete = false;
                references.Declarations.Record(attachment, "$", null);
                continue;
            }
            budget.AddDrawableReferences(1);
            IWorkArchiveRecord? drawable = references.ReadOne(attachment, payload, 1);
            // The qualified inline mode uses explicit zero offsets. Other modes stay unresolved.
            bool placement = true;
            foreach (int field in new[] { 2, 3, 4, 5 }) {
                bool integer = field == 2 || field == 4;
                bool valid = payload.FieldCount(field) == 1
                    && !payload.HasUnexpectedWireKind(field, integer ? IWorkWireKind.Varint : IWorkWireKind.Fixed32)
                    && (integer ? payload.GetUnsigned(field) == 0 : payload.GetFloat(field) == 0f);
                if (!valid) {
                    placement = false;
                    references.Declarations.Record(attachment, field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        payload.FieldCount(field), IWorkSourceDeclarationIssueKind.InvalidValue);
                }
            }
            if (payload.TotalFieldCount != 5) {
                complete = false;
                references.Declarations.Record(attachment, "$", payload.TotalFieldCount, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
                continue;
            }
            if (payload.FieldCount(1) != 1 || payload.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
                || !placement || drawable?.MessageType is not (3005 or 6000 or 6007)) {
                complete = false;
                continue;
            }
            result.Add(boundary.Index, new IWorkInlineObject(boundary.Index,
                new IWorkObjectIdentity(attachment), new IWorkObjectIdentity(drawable)));
        }
        return result;
    }
}
