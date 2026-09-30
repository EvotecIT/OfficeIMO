namespace OfficeIMO.IWork.Internal;

internal static class IWorkSourceUnitInventory {
    internal static IReadOnlyList<IWorkSourceUnit> Create(IWorkSourceDocument source, IWorkProjectionKind projectionKind,
        IEnumerable<IWorkObjectIdentity?>? reconstructed, IEnumerable<IWorkObjectIdentity?>? omitted) {
        var recoveredIds = Identifiers(reconstructed);
        var omittedIds = Identifiers(omitted);
        var result = new List<IWorkSourceUnit>();
        foreach (IWorkArchiveRecord record in source.Records) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (!record.IsPrimary || !recoveredIds.Contains(record.Identifier) && !omittedIds.Contains(record.Identifier)
                || Classify(source.Kind, record.MessageType) is not { } kind) continue;
            IWorkSourceUnitDisposition disposition = IWorkSourceUnitDisposition.Unassessed;
            if (projectionKind == IWorkProjectionKind.EditableReconstruction) {
                if (recoveredIds.Contains(record.Identifier)) disposition = IWorkSourceUnitDisposition.Reconstructed;
                else if (omittedIds.Contains(record.Identifier)) disposition = IWorkSourceUnitDisposition.Omitted;
            }
            result.Add(new IWorkSourceUnit(kind, new IWorkObjectIdentity(record), disposition));
        }
        return result;

        HashSet<ulong> Identifiers(IEnumerable<IWorkObjectIdentity?>? identities) {
            var result = new HashSet<ulong>();
            if (identities is null) return result;
            foreach (IWorkObjectIdentity? identity in identities) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (identity is not null) result.Add(identity.RecordIdentifier);
            }
            return result;
        }
    }

    private static IWorkSourceUnitKind? Classify(IWorkDocumentKind documentKind, uint type) => type switch {
        10000 when documentKind == IWorkDocumentKind.Pages => IWorkSourceUnitKind.Document,
        1 when documentKind is IWorkDocumentKind.Numbers or IWorkDocumentKind.Keynote => IWorkSourceUnitKind.Document,
        2 when documentKind == IWorkDocumentKind.Numbers => IWorkSourceUnitKind.Sheet,
        5 when documentKind == IWorkDocumentKind.Keynote => IWorkSourceUnitKind.Slide,
        6000 or 6007 => IWorkSourceUnitKind.Table,
        2001 => IWorkSourceUnitKind.Text,
        3005 => IWorkSourceUnitKind.Image,
        _ => null
    };
}
