using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkPagesReader {
    private static bool ReadSectionSelectionFlag(IWorkWireMessage message, IWorkArchiveRecord owner,
        int field, IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics,
        ref bool complete) {
        if (!message.HasField(field)) return false;
        ulong? value = message.GetUnsigned(field);
        if (message.FieldCount(field) == 1
            && !message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
            && value.HasValue && value.Value <= 1) return value.Value == 1;
        complete = false;
        references.Declarations.Record(owner, field.ToString(System.Globalization.CultureInfo.InvariantCulture),
            message.FieldCount(field), IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
            "A Pages header/footer selection flag is malformed; editable reconstruction is incomplete.",
            owner.EntryPath, owner.Identifier));
        return false;
    }
}
