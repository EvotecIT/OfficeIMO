namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static IReadOnlyDictionary<uint, string> ReadStrings(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkArchiveRecord model,
        IWorkSourceReferenceIssueCollector references, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries,
        out bool fullyReconstructed) {
        var strings = new Dictionary<uint, string>();
        fullyReconstructed = true;
        IWorkArchiveRecord? list = references.ReadOne(model, store, 4, "4/4");
        if (list == null) {
            fullyReconstructed = !store.HasField(4);
            return strings;
        }
        if (!TryGetCatalogEntryCount(list, maximumEntries, options, "string",
                out int entryCount)) {
            fullyReconstructed = false;
            return strings;
        }
        projectionBudget.AddTableCatalogEntries(entryCount);
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            index.Message(list), 3, out bool malformedEntries);
        if (malformedEntries) fullyReconstructed = false;
        foreach (IWorkWireMessage entry in entries) {
            ulong? key = entry.GetUnsigned(1);
            string? value = entry.GetString(3);
            if (entry.FieldCount(1) != 1
                || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !key.HasValue || key.Value > uint.MaxValue || value == null) {
                fullyReconstructed = false;
                continue;
            }
            projectionBudget.AddTextCharacters(value.Length);
            uint normalizedKey = (uint)key.Value;
            if (strings.ContainsKey(normalizedKey)) fullyReconstructed = false;
            else strings.Add(normalizedKey, value);
        }
        return strings;
    }

    private static IReadOnlyDictionary<uint, IWorkWireMessage> ReadFormulas(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkArchiveRecord model,
        IWorkSourceReferenceIssueCollector references, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries,
        out bool fullyReconstructed, out bool catalogEnvelopeComplete) {
        var formulas = new Dictionary<uint, IWorkWireMessage>();
        var ambiguousIdentifiers = new HashSet<uint>();
        fullyReconstructed = true;
        catalogEnvelopeComplete = true;
        IWorkArchiveRecord? list = references.ReadOne(model, store, 6, "4/6");
        if (list == null) {
            fullyReconstructed = catalogEnvelopeComplete = !store.HasField(6);
            return formulas;
        }
        if (!TryGetCatalogEntryCount(list, maximumEntries, options, "formula",
                out int entryCount)) {
            fullyReconstructed = false;
            catalogEnvelopeComplete = false;
            return formulas;
        }
        projectionBudget.AddTableCatalogEntries(entryCount);
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            index.Message(list), 3, out bool malformedEntries);
        if (malformedEntries) fullyReconstructed = false;
        foreach (IWorkWireMessage entry in entries) {
            ulong? key = entry.GetUnsigned(1);
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(entry, 5, out bool malformedFormula);
            if (entry.FieldCount(1) != 1
                || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !key.HasValue || key.Value > uint.MaxValue || malformedFormula || formula == null) {
                fullyReconstructed = false;
                continue;
            }
            uint normalizedKey = (uint)key.Value;
            if (ambiguousIdentifiers.Contains(normalizedKey)) {
                fullyReconstructed = false;
            } else if (formulas.ContainsKey(normalizedKey)) {
                formulas.Remove(normalizedKey);
                ambiguousIdentifiers.Add(normalizedKey);
                fullyReconstructed = false;
            } else {
                formulas.Add(normalizedKey, formula);
            }
        }
        return formulas;
    }

    internal static bool TryGetCatalogEntryCount(IWorkArchiveRecord list, int maximumEntries,
        IWorkReadOptions options, string catalogName, out int declaredEntryCount) {
        declaredEntryCount = 0;
        int totalFieldCount;
        int identifierFieldCount;
        int metadataFieldCount;
        try {
            declaredEntryCount = IWorkProtobuf.CountFields(
                list.Payload, 3, options.MaximumProtobufFieldCount,
                out totalFieldCount);
            identifierFieldCount = IWorkProtobuf.CountFields(
                list.Payload, 1, options.MaximumProtobufFieldCount);
            metadataFieldCount = IWorkProtobuf.CountFields(
                list.Payload, 2, options.MaximumProtobufFieldCount);
        } catch (InvalidDataException exception)
            when (!IWorkProtobuf.IsFieldLimitException(exception)) {
            declaredEntryCount = 0;
            return false;
        }
        if (identifierFieldCount > 1 || metadataFieldCount > 1
            || totalFieldCount - declaredEntryCount
                != identifierFieldCount + metadataFieldCount) {
            declaredEntryCount = 0;
            return false;
        }
        if (declaredEntryCount > maximumEntries) {
            throw new InvalidDataException(
                $"An iWork {catalogName} catalog exceeds the remaining table-catalog limit of {maximumEntries}.");
        }
        return true;
    }

}
