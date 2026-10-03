using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkKeynoteReader {
    private static void AssessBuildDeclarations(IWorkWireMessage message, IWorkArchiveRecord slide,
        IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics,
        ref bool complete) {
        bool declared = false;
        // KN.SlideArchive: builds, inline buildChunkArchives, referenced buildChunks.
        // Retain the envelope without interpreting unsupported animation content.
        foreach (int field in new[] { 2, 3, 43 }) {
            if (!message.HasField(field)) continue;
            declared = true;
            references.Declarations.Record(slide,
                field.ToString(System.Globalization.CultureInfo.InvariantCulture), message.FieldCount(field),
                message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)
                    ? IWorkSourceDeclarationIssueKind.MalformedMessage
                    : IWorkSourceDeclarationIssueKind.UnsupportedField);
        }
        if (!declared) return;
        complete = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_BUILDS_UNSUPPORTED",
            "A selected Keynote slide declares builds or animation chunks that are not reconstructed; editable conversion is incomplete.",
            slide.EntryPath, slide.Identifier));
    }
}
