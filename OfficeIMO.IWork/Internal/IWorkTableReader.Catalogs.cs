namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static IReadOnlyDictionary<uint, string> ReadStrings(IWorkSourceDocument source,
        IWorkWireMessage store, IWorkArchiveRecord model,
        IWorkSourceReferenceIssueCollector references, IWorkProjectionBudget projectionBudget,
        out bool fullyReconstructed) {
        var strings = new Dictionary<uint, string>();
        fullyReconstructed = true;
        IWorkArchiveRecord? list = references.ReadOne(model, store, 4, "4/4");
        if (list == null) {
            fullyReconstructed = !store.HasField(4);
            return strings;
        }
        IWorkTableCatalogIndex catalog = IWorkTableCatalogIndex.Read(source, list, projectionBudget, references, "string");
        fullyReconstructed = catalog.IsComplete;
        foreach (var entry in catalog.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            string? value = entry.Message.GetString(3);
            if (value == null) {
                fullyReconstructed = false;
                references.Declarations.Record(list, IWorkTableCatalogIndex.EntryPath(entry.Position) + "/3",
                    entry.Message.FieldCount(3), IWorkSourceDeclarationIssueKind.InvalidValue);
                continue;
            }
            projectionBudget.AddTextCharacters(value.Length);
            if (catalog.CanResolveKey(entry.Key)) strings.Add(entry.Key, value);
        }
        return strings;
    }

    private static IReadOnlyDictionary<uint, IWorkWireMessage> ReadFormulas(IWorkSourceDocument source,
        IWorkWireMessage store, IWorkArchiveRecord model,
        IWorkSourceReferenceIssueCollector references, IWorkProjectionBudget projectionBudget,
        out bool fullyReconstructed, out bool catalogEnvelopeComplete) {
        var formulas = new Dictionary<uint, IWorkWireMessage>();
        fullyReconstructed = true;
        catalogEnvelopeComplete = true;
        IWorkArchiveRecord? list = references.ReadOne(model, store, 6, "4/6");
        if (list == null) {
            fullyReconstructed = catalogEnvelopeComplete = !store.HasField(6);
            return formulas;
        }
        IWorkTableCatalogIndex catalog = IWorkTableCatalogIndex.Read(source, list, projectionBudget, references, "formula");
        fullyReconstructed = catalog.IsComplete;
        catalogEnvelopeComplete = catalog.EnvelopeIsComplete;
        foreach (var entry in catalog.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(entry.Message, 5, out bool malformed);
            if (malformed || formula == null) {
                fullyReconstructed = false;
                references.Declarations.Record(list, IWorkTableCatalogIndex.EntryPath(entry.Position) + "/5",
                    entry.Message.FieldCount(5), entry.Message.FieldCount(5) != 1
                        || entry.Message.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
                        ? IWorkSourceDeclarationIssueKind.RejectedMessageSet
                        : IWorkSourceDeclarationIssueKind.MalformedMessage);
                continue;
            }
            if (catalog.CanResolveKey(entry.Key)) formulas.Add(entry.Key, formula);
        }
        return formulas;
    }
}
