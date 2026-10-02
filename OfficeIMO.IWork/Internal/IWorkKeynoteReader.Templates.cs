using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkKeynoteReader {
    private static void AssessTemplateReference(IWorkSourceDocument source, IWorkArchiveRecord slide, IWorkWireMessage message,
        IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics, ref bool complete) {
        if (!message.HasField(17)) return;
        IWorkArchiveRecord? template = references.ReadOne(slide, message, 17, allowedType: static type => type == 5);
        if (message.FieldCount(17) == 1 && !message.HasUnexpectedWireKind(17, IWorkWireKind.Bytes)
            && template?.MessageType == 5) {
            if (template.Identifier == slide.Identifier) {
                references.Declarations.Record(slide, "17", 1, IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            } else {
                try {
                    source.Index.Message(template);
                    return;
                } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                    references.Declarations.Record(template, "$", null);
                }
            }
        }
        complete = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_TEMPLATE_UNRESOLVED",
            "A selected Keynote slide has an unresolved template reference; editable reconstruction is incomplete.",
            slide.EntryPath, slide.Identifier));
    }
}
