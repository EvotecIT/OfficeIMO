namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static bool? ReadTableStyleSettings(IWorkObjectIndex index, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction, out bool fillDefaultsSupported, out IWorkCellFill? bandedBodyFill) {
        fillDefaultsSupported = true;
        bandedBodyFill = null;
        if (!message.HasField(3)) return null;
        IWorkArchiveRecord? style = references.ReadOne(model, message, 3);
        bool complete = message.FieldCount(3) == 1
            && !message.HasUnexpectedWireKind(3, IWorkWireKind.Bytes) && style?.MessageType == 6003;
        bool? autoResizeRows = null;
        bool bandingComplete = true;
        bool banded = false;
        IWorkArchiveRecord? bandOwner = null;
        IWorkWireMessage? bandProperties = null;
        if (complete) {
            var chain = IWorkStyleReader.ReadChain(index, style!.Identifier,
                budget.MaximumTextStyleInheritanceDepth, type => type == 6003,
                tolerateStyleDepth: true, references, ref complete);
            for (int i = chain.Count - 1; i >= 0; i--) {
                var (owner, payload) = chain[i];
                IWorkWireMessage? properties = IWorkObjectIndex.TryGetMessage(payload, 11, out bool malformed);
                if (malformed || payload.FieldCount(11) > 1
                    || payload.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                    || payload.HasField(11) && properties == null) {
                    references.Declarations.Record(owner, "11", payload.FieldCount(11));
                    complete = false;
                    continue;
                }
                if (properties == null) continue;
                if (properties.HasField(1)) {
                    ulong? banding = properties.GetUnsigned(1);
                    if (properties.FieldCount(1) != 1 || properties.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                        || banding is not (0 or 1)) {
                        bandingComplete = false;
                        references.Declarations.Record(owner, "11/1", properties.FieldCount(1),
                            IWorkSourceDeclarationIssueKind.InvalidValue);
                    } else banded = banding == 1;
                }
                if (properties.HasField(2)) { bandOwner = owner; bandProperties = properties; }
                if (!properties.HasField(22)) continue;
                ulong? value = properties.GetUnsigned(22);
                if (properties.FieldCount(22) != 1
                    || properties.HasUnexpectedWireKind(22, IWorkWireKind.Varint)
                    || value is not (0 or 1)) {
                    references.Declarations.Record(owner, "11/22", properties.FieldCount(22),
                        IWorkSourceDeclarationIssueKind.InvalidValue);
                    complete = false;
                    continue;
                }
                autoResizeRows = value == 1;
            }
        }
        // Decode the effective band fill only when banding is active. An inactive inherited
        // gradient or newer color model has no effect on the reconstructed table.
        if (banded) {
            if (bandProperties == null) bandingComplete = false;
            else IWorkCellFillReader.Read(bandOwner!, bandProperties, references,
                ref bandedBodyFill, ref bandingComplete, field: 2);
        }
        fillDefaultsSupported = complete && bandingComplete;
        if (!fillDefaultsSupported) bandedBodyFill = null;
        if (complete) return autoResizeRows;
        references.Declarations.Record(model, "3", message.FieldCount(3),
            IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        supportsEditableReconstruction = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_TABLE_ROW_SIZING_UNSUPPORTED",
            "An iWork table's automatic row sizing is unresolved or malformed; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
        return null;
    }
}
