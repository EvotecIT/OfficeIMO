using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkPagesReader {
    private static IReadOnlyList<IWorkArchiveRecord> CollectDocumentDrawables(IWorkObjectIndex index,
        IWorkArchiveRecord document, IWorkWireMessage documentMessage,
        IWorkProjectionBudget projectionBudget, IWorkSourceReferenceIssueCollector references,
        out IReadOnlyDictionary<ulong, int> pageIndexes, out bool complete) {
        complete = true;
        var identifiers = new HashSet<ulong>();
        var ordered = new List<IWorkArchiveRecord>();
        var pages = new Dictionary<ulong, int>();
        void Add(IWorkArchiveRecord record) {
            if (identifiers.Add(record.Identifier)) ordered.Add(record);
        }

        IWorkArchiveRecord? zOrder = references.ReadOne(document, documentMessage, 20);
        if (documentMessage.HasUnexpectedWireKind(20, IWorkWireKind.Bytes)
            || documentMessage.HasField(20) && zOrder == null) complete = false;
        if (zOrder != null) {
            int zOrderReferenceCount = 0;
            IWorkWireMessage? zOrderMessage = null;
            try {
                zOrderReferenceCount = IWorkProtobuf.CountFields(
                    zOrder.Payload, 1, projectionBudget.MaximumProtobufFieldCount);
                if (!TryReadMessage(index, zOrder, references, out zOrderMessage)) complete = false;
            } catch (InvalidDataException exception)
                when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(zOrder, "$", null);
                complete = false;
            }
            if (zOrderMessage != null) {
                projectionBudget.AddDrawableReferences(zOrderReferenceCount);
                int unresolvedZOrderCount;
                var zOrderOccurrences = new HashSet<ulong>();
                foreach (IWorkArchiveRecord record in references.ReadAll(
                             zOrder, zOrderMessage, 1, out unresolvedZOrderCount)) {
                    if (!zOrderOccurrences.Add(record.Identifier)) complete = false;
                    Add(record);
                }
                if (unresolvedZOrderCount > 0) complete = false;
            }
        }
        IWorkArchiveRecord? floating = references.ReadOne(document, documentMessage, 3);
        if (documentMessage.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
            || documentMessage.HasField(3) && floating == null) complete = false;
        if (floating != null) {
            IReadOnlyList<IWorkWireMessage> pageGroups;
            int pageGroupCount = 0;
            try {
                pageGroupCount = IWorkProtobuf.CountFields(floating.Payload, 1,
                    projectionBudget.MaximumProtobufFieldCount, out int totalFieldCount);
                if (totalFieldCount != pageGroupCount
                    || !TryReadMessage(index, floating, references, out IWorkWireMessage floatingMessage)) {
                    complete = false;
                    references.Declarations.Record(floating, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
                    pageGroups = Array.Empty<IWorkWireMessage>();
                } else {
                    pageGroups = IWorkObjectIndex.TryGetMessages(floatingMessage, 1,
                        out bool malformedPageGroups);
                    if (malformedPageGroups) {
                        complete = false;
                        references.Declarations.Record(floating, "1", floatingMessage.FieldCount(1),
                            IWorkSourceDeclarationIssueKind.RejectedMessageSet);
                    }
                }
            } catch (InvalidDataException exception)
                when (!IWorkProtobuf.IsLimitException(exception)) {
                complete = false;
                references.Declarations.Record(floating, "$", null);
                pageGroups = Array.Empty<IWorkWireMessage>();
            }
            projectionBudget.AddDrawableReferences(pageGroupCount);
            for (int pageGroupIndex = 0; pageGroupIndex < pageGroups.Count; pageGroupIndex++) {
                IWorkWireMessage pageGroup = pageGroups[pageGroupIndex];
                foreach (int field in new[] { 2, 3, 4 }) {
                    projectionBudget.AddDrawableReferences(pageGroup.FieldCount(field));
                    var fieldOccurrences = new HashSet<ulong>();
                    IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
                        pageGroup, field, out bool malformedEntries);
                    if (malformedEntries) {
                        complete = false;
                        references.Declarations.Record(floating,
                            FormattableString.Invariant($"1[{pageGroupIndex + 1}]/{field}"),
                            pageGroup.FieldCount(field), IWorkSourceDeclarationIssueKind.RejectedMessageSet);
                    }
                    int entryIndex = 0;
                    foreach (IWorkWireMessage entry in entries) {
                        entryIndex++;
                        IWorkArchiveRecord? record = references.ReadOne(floating, entry, 1,
                            FormattableString.Invariant($"1[{pageGroupIndex + 1}]/{field}[{entryIndex}]/1"));
                        if (entry.FieldCount(1) != 1
                            || entry.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
                            || record == null) complete = false;
                        else if (record != null) {
                            if (!fieldOccurrences.Add(record.Identifier)) complete = false;
                            if (pages.TryGetValue(record.Identifier, out int existingPageIndex)
                                && existingPageIndex != pageGroupIndex + 1) complete = false;
                            else pages[record.Identifier] = pageGroupIndex + 1;
                            Add(record);
                        }
                    }
                }
            }
        }
        var reachable = new HashSet<ulong>(index.ReachableFrom(document).Select(record => record.Identifier));
        foreach (IWorkArchiveRecord record in index.PrimaryRecords.Where(record =>
                     reachable.Contains(record.Identifier)
                     && record.MessageType is ShapeInfoArchive or 3005 or 6000 or 6007)) Add(record);
        pageIndexes = pages;
        return Array.AsReadOnly(ordered.ToArray());
    }

}
