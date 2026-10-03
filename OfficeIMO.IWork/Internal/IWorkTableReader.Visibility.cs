using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static readonly int[] VisibilityCountFields = { 14, 15, 40, 41, 42 };

    private static void AssessVisibilityCounts(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, int rows, int columns, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        bool complete = true;
        // TableModelArchive declares total hidden rows/columns (14/15), filtered rows (40),
        // and user-hidden rows/columns (41/42) independently of the dimension headers.
        // These counts establish a fidelity gap, not the identities of hidden cells.
        foreach (int field in VisibilityCountFields) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (!message.HasField(field)) continue;
            ulong? count = message.GetUnsigned(field);
            int dimensionCount = field == 15 || field == 42 ? columns : rows;
            bool invalid = message.FieldCount(field) != 1
                || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
                || count == null || count.Value > (ulong)dimensionCount;
            if (!invalid && count == 0) continue;
            complete = false;
            references.Declarations.Record(model, field.ToString(CultureInfo.InvariantCulture),
                message.FieldCount(field), invalid ? IWorkSourceDeclarationIssueKind.InvalidValue
                    : IWorkSourceDeclarationIssueKind.UnsupportedField);
        }
        if (complete) return;
        supportsEditableReconstruction = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_VISIBILITY_UNASSESSED",
            "An iWork table declares hidden or filtered rows/columns, or unreadable visibility counts. "
                + "Visibility is not reconstructed; partial editable output and Reader may include content hidden in the source.",
            model.EntryPath, model.Identifier, global::OfficeIMO.OfficeConversionLossKind.Unassessed));
    }
}
