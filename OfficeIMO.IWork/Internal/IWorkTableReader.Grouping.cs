namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    // Category and pivot layouts can introduce visual summary rows and reorder base rows.
    // Disabled category metadata occurs on ordinary native tables and is safe to leave opaque.
    private static bool AssessGrouping(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        bool ordinary = true;
        if (message.HasField(85)) {
            references.ReadOne(model, message, 85, allowedType: type => type == 6370);
            references.Declarations.Record(model, "85", message.FieldCount(85), IWorkSourceDeclarationIssueKind.UnsupportedField);
            ordinary = false;
        }
        try {
            if (message.HasField(86)) {
                budget.AddTableDimensionEntries(message.FieldCount(86));
                var owner = references.ReadOne(model, message, 86, allowedType: type => type == 6372);
                if (owner?.MessageType != 6372 || message.FieldCount(86) != 1
                    || message.HasUnexpectedWireKind(86, IWorkWireKind.Bytes)) Reject(model, "86", message.FieldCount(86));
                else {
                    var payload = source.Index.Message(owner);
                    budget.AddTableDimensionEntries(payload.FieldCount(1));
                    var groups = references.ReadAll(owner, payload, 1, out int unresolved, allowedType: type => type == 6373);
                    if (unresolved > 0 || groups.Count != payload.FieldCount(1)) Reject(owner, "1", payload.FieldCount(1));
                    foreach (var group in groups) {
                        source.CancellationToken.ThrowIfCancellationRequested();
                        if (group.MessageType != 6373) Reject(owner, "1", payload.FieldCount(1));
                        else AssessGroup(group, source.Index.Message(group), "6");
                    }
                }
            } else if (message.HasField(81)) {
                var owner = IWorkObjectIndex.TryGetMessage(message, 81, out bool malformed);
                if (malformed || owner == null || message.FieldCount(81) != 1
                    || message.HasUnexpectedWireKind(81, IWorkWireKind.Bytes)) Reject(model, "81", message.FieldCount(81));
                else {
                    budget.AddTableDimensionEntries(owner.FieldCount(2));
                    var groups = IWorkObjectIndex.TryGetMessages(owner, 2, out malformed);
                    if (malformed || groups.Count != owner.FieldCount(2)) Reject(model, "81/2", owner.FieldCount(2));
                    for (int index = 0; index < groups.Count; index++) {
                        source.CancellationToken.ThrowIfCancellationRequested();
                        AssessGroup(model, groups[index], "81/2[" + (index + 1) + "]/6");
                    }
                }
            }
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            // Configured bounds are fatal; malformed selected metadata is retained as incomplete.
            Reject(model, message.HasField(86) ? "86" : "81", null);
        }
        if (!ordinary) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_GROUPING_UNSUPPORTED",
                "Active or unresolved category/pivot layouts can change displayed row order, summary rows and borders; base cell values remain recoverable.",
                model.EntryPath, model.Identifier, global::OfficeIMO.OfficeConversionLossKind.Unassessed));
        }
        return ordinary;

        void AssessGroup(IWorkArchiveRecord owner, IWorkWireMessage payload, string path) {
            bool valid = payload.FieldCount(6) == 1 && !payload.HasUnexpectedWireKind(6, IWorkWireKind.Varint)
                && payload.GetUnsigned(6) is 0 or 1;
            if (!valid) Reject(owner, path, payload.FieldCount(6));
            else if (payload.GetUnsigned(6) == 1) {
                ordinary = false; references.Declarations.Record(owner, path, 1, IWorkSourceDeclarationIssueKind.UnsupportedField);
            }
        }

        void Reject(IWorkArchiveRecord owner, string path, int? count) {
            ordinary = false; references.Declarations.Record(owner, path, count, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        }
    }
}
