using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private const uint FilterSetArchive = 6220;
    private static readonly int[] UnassessedExtentFields = { 5, 7, 9, 10, 11 };

    private static void AssessHiddenStates(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, int rows, int columns, HashSet<int> hiddenRows, HashSet<int> hiddenColumns,
        IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        bool complete = true;
        AssessVisibilityFilter(source, model, message, 38, "38", budget, references, ref complete);
        if (message.HasField(70)) {
            budget.AddTableDimensionEntries(message.FieldCount(70));
            IWorkWireMessage? owner = ReadVisibilityMessage(model, message, 70, "70", references, ref complete);
            if (owner != null) {
                budget.AddTableDimensionEntries(owner.FieldCount(2));
                int position = 0;
                foreach (IWorkWireValue value in owner.EnumerateValues(2)) {
                    source.CancellationToken.ThrowIfCancellationRequested();
                    string path = "70/2[" + (++position).ToString(CultureInfo.InvariantCulture) + "]";
                    IWorkWireMessage? states = ReadVisibilityEntry(model, owner, value, path, references, ref complete);
                    if (states == null) continue;
                    foreach (int axis in new[] { 2, 3 }) {
                        source.CancellationToken.ThrowIfCancellationRequested();
                        string extentPath = path + "/" + axis.ToString(CultureInfo.InvariantCulture);
                        budget.AddTableDimensionEntries(states.FieldCount(axis));
                        IWorkWireMessage? extent = ReadVisibilityMessage(model, states, axis, extentPath,
                            references, ref complete, required: true);
                        if (extent == null) continue;
                        AssessHiddenExtent(source, model, message, extent, axis == 2 ? 0ul : 1ul, extentPath,
                            axis == 2 ? columns : rows, owner.FieldCount(2) == 1, axis == 2 ? hiddenColumns : hiddenRows,
                            budget, references, ref complete);
                    }
                }
            }
        }
        if (complete) return;
        supportsEditableReconstruction = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_HIDDEN_STATES_UNASSESSED",
            "An iWork table selects hidden/filtered states, collapsed groups or active filters, or contains unreadable visibility declarations. "
                + "Visibility is not reconstructed; partial editable output and Reader may include source-hidden content.",
            model.EntryPath, model.Identifier, global::OfficeIMO.OfficeConversionLossKind.Unassessed));
    }

    private static void AssessHiddenExtent(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage modelMessage, IWorkWireMessage extent, ulong expectedDirection, string path,
        int dimensionCount, bool singleStateSet, HashSet<int> hidden, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        bool directionValid = extent.FieldCount(3) == 1 && !extent.HasUnexpectedWireKind(3, IWorkWireKind.Varint)
            && extent.GetUnsigned(3) == expectedDirection;
        if (!directionValid) {
            complete = false;
            references.Declarations.Record(model, path + "/3", extent.FieldCount(3), IWorkSourceDeclarationIssueKind.InvalidValue);
        }
        AssessVisibilityFlag(model, extent, 6, path + "/6", references, ref complete);
        foreach (int field in UnassessedExtentFields) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (!extent.HasField(field)) continue;
            budget.AddTableDimensionEntries(extent.FieldCount(field));
            complete = false;
            references.Declarations.Record(model, path + "/" + field.ToString(CultureInfo.InvariantCulture),
                extent.FieldCount(field), IWorkSourceDeclarationIssueKind.UnsupportedField);
        }
        foreach (int field in new[] { 2, 12 }) {
            budget.AddTableDimensionEntries(extent.FieldCount(field));
            int position = 0;
            bool entriesComplete = true;
            var baseStates = new List<(IWorkWireMessage Message, string Path)>();
            foreach (IWorkWireValue value in extent.EnumerateValues(field)) {
                source.CancellationToken.ThrowIfCancellationRequested();
                string entryPath = path + "/" + field.ToString(CultureInfo.InvariantCulture)
                    + "[" + (++position).ToString(CultureInfo.InvariantCulture) + "]";
                IWorkWireMessage? state = ReadVisibilityEntry(model, extent, value, entryPath, references, ref complete);
                if (state == null) { entriesComplete = false; continue; }
                if (field == 2) baseStates.Add((state, entryPath));
                foreach (int flag in field == 2 ? new[] { 3, 4 } : new[] { 2, 3, 4 })
                    AssessVisibilityFlag(model, state, flag, entryPath + "/" + flag.ToString(CultureInfo.InvariantCulture),
                        references, ref complete);
            }
            if (field != 2) continue;
            bool selected = baseStates.Any(state => state.Message.FieldCount(2) == 1
                && !state.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint) && state.Message.GetUnsigned(2) == 1);
            IReadOnlyList<int> recovered = Array.Empty<int>();
            bool resolved = selected && entriesComplete && directionValid && singleStateSet
                && IWorkTableVisibilityMap.TryResolve(source, model, modelMessage, dimensionCount, expectedDirection == 0,
                    baseStates, budget, references, out recovered);
            if (resolved) {
                foreach (int index in recovered!) hidden.Add(index);
            } else {
                foreach (var state in baseStates)
                    AssessVisibilityFlag(model, state.Message, 2, state.Path + "/2", references, ref complete);
            }
        }
        AssessVisibilityFilter(source, model, extent, 8, path + "/8", budget, references, ref complete);
    }

    private static void AssessVisibilityFlag(IWorkArchiveRecord model, IWorkWireMessage message, int field,
        string path, IWorkSourceReferenceIssueCollector references, ref bool complete, bool required = false) {
        if (!required && !message.HasField(field)) return;
        ulong? value = message.GetUnsigned(field);
        bool invalid = message.FieldCount(field) != 1 || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
            || value == null || value > 1;
        if (!invalid && value == 0) return;
        complete = false;
        references.Declarations.Record(model, path, message.FieldCount(field), invalid
            ? IWorkSourceDeclarationIssueKind.InvalidValue : IWorkSourceDeclarationIssueKind.UnsupportedField);
    }

    private static void AssessVisibilityFilter(IWorkSourceDocument source, IWorkArchiveRecord owner,
        IWorkWireMessage message, int field, string path, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        if (!message.HasField(field)) return;
        budget.AddTableDimensionEntries(message.FieldCount(field));
        IWorkArchiveRecord? filter = references.ReadOne(owner, message, field, path);
        if (filter == null) { complete = false; return; }
        if (filter.MessageType != FilterSetArchive) {
            complete = false;
            references.Declarations.Record(filter, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return;
        }
        IWorkWireMessage filterMessage;
        try { filterMessage = source.Index.Message(filter); }
        catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            complete = false;
            references.Declarations.Record(filter, "$", null);
            return;
        }
        // Disabled filters do not select their rules. Do not traverse inactive rule/formula graphs.
        AssessVisibilityFlag(filter, filterMessage, 2, "2", references, ref complete, required: true);
    }

    private static IWorkWireMessage? ReadVisibilityMessage(IWorkArchiveRecord owner, IWorkWireMessage parent,
        int field, string path, IWorkSourceReferenceIssueCollector references, ref bool complete, bool required = false) {
        if (!required && !parent.HasField(field)) return null;
        IWorkWireMessage? message = IWorkObjectIndex.TryGetMessage(parent, field, out bool malformed);
        if (!malformed && message != null) return message;
        complete = false;
        references.Declarations.Record(owner, path, parent.FieldCount(field), parent.FieldCount(field) > 1
            ? IWorkSourceDeclarationIssueKind.RejectedMessageSet : IWorkSourceDeclarationIssueKind.MalformedMessage);
        return null;
    }

    private static IWorkWireMessage? ReadVisibilityEntry(IWorkArchiveRecord owner, IWorkWireMessage parent,
        IWorkWireValue value, string path, IWorkSourceReferenceIssueCollector references, ref bool complete) {
        if (value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
            try { return parent.ParseNestedMessage(value.Bytes); }
            catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { }
        }
        complete = false;
        references.Declarations.Record(owner, path, 1);
        return null;
    }
}
