using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkKeynoteReader {
    private static IWorkCellFill? ReadBackground(IWorkSourceDocument source, IWorkArchiveRecord slide,
        IWorkWireMessage message, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool complete) {
        if (!message.HasField(1) && !message.HasField(17)) return null;
        IWorkArchiveRecord? style = references.ReadOne(slide, message, 1, allowedType: static type => type == 9);
        bool supported = message.FieldCount(1) == 1 && !message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            && style?.MessageType == 9;
        IWorkCellFill? fill = null;
        if (supported) {
            var chain = IWorkStyleReader.ReadChain(source.Index, style!.Identifier,
                budget.MaximumTextStyleInheritanceDepth, static type => type == 9,
                tolerateStyleDepth: true, references, ref supported);
            for (int i = chain.Count - 1; i >= 0; i--) {
                source.CancellationToken.ThrowIfCancellationRequested();
                var (owner, payload) = chain[i];
                IWorkWireMessage? properties = IWorkObjectIndex.TryGetMessage(payload, 11, out bool malformed);
                if (malformed || payload.FieldCount(11) > 1 || payload.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                    || payload.HasField(11) && properties == null) {
                    references.Declarations.Record(owner, "11", payload.FieldCount(11));
                    supported = false;
                    continue;
                }
                if (properties != null) IWorkCellFillReader.Read(owner, properties, references, ref fill, ref supported);
            }
        }
        if (message.HasField(17) && fill == null && (supported || !message.HasField(1))) {
            references.Declarations.Record(slide, "17", message.FieldCount(17),
                IWorkSourceDeclarationIssueKind.UnsupportedField);
            supported = false;
        }
        if (supported) return fill;
        complete = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_BACKGROUND_UNSUPPORTED",
            "A selected Keynote slide background has an unresolved style, unsupported fill or unqualified template background; editable reconstruction is incomplete.",
            slide.EntryPath, slide.Identifier));
        return null;
    }
}
