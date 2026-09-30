using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static IReadOnlyList<IWorkTableMergeRange> ReadMergedRanges(IWorkSourceDocument source,
        IWorkWireMessage table, int rowCount, int columnCount, int maximumRanges, int maximumFormulaNodes,
        IWorkArchiveRecord model, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        IWorkWireMessage? mergeOwner = IWorkObjectIndex.TryGetMessage(table, 47, out bool malformedOwner);
        if (malformedOwner) {
            RecordMergeMessageFailure(model, table, 47, "47", references);
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (mergeOwner == null) return Array.Empty<IWorkTableMergeRange>();
        byte[]? formulaStoreBytes = mergeOwner.GetBytes(2);
        int pairCount;
        try {
            pairCount = formulaStoreBytes == null || mergeOwner.FieldCount(2) != 1
                || mergeOwner.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                    ? -1 : mergeOwner.CountNestedFields(formulaStoreBytes, 3);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            pairCount = -1;
        }
        if (pairCount > maximumRanges)
            throw new InvalidDataException($"iWork table merged-range count exceeds the configured limit of {maximumRanges} in object {model.Identifier}.");
        if (pairCount < 0) {
            RecordMergeMessageFailure(model, mergeOwner, 2, "47/2", references);
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        IWorkWireMessage formulaStore = mergeOwner.ParseNestedMessage(formulaStoreBytes!);
        var declarations = new List<MergeDeclaration>();
        var knownRanges = new Dictionary<(int, int, int, int), int>();
        bool unknownRange = false;
        int position = 0;
        foreach (IWorkWireValue value in formulaStore.EnumerateValues(3)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            string path = MergePairPath(++position);
            IWorkWireMessage? pair = null;
            if (value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
                try { pair = formulaStore.ParseNestedMessage(value.Bytes); }
                catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { }
            }
            if (pair == null) {
                references.Declarations.Record(model, path, 1);
                unknownRange = true;
                MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(pair, 2, out bool malformedFormula);
            if (malformedFormula || formula == null) {
                RecordMergeMessageFailure(model, pair, 2, path + "/2", references);
                unknownRange = true;
                MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            if (!IWorkFormulaReader.TryReadAbsoluteRange(formula, maximumFormulaNodes,
                    out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                references.Declarations.Record(model, path + "/2", 1, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
                unknownRange = true;
                MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            bool inBounds = lastRow < rowCount && lastColumn < columnCount;
            if (!inBounds) {
                references.Declarations.Record(model, path + "/2", 1, IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            }
            if (firstRow >= rowCount || firstColumn >= columnCount
                || inBounds && firstRow == lastRow && firstColumn == lastColumn) continue;
            var key = (firstRow, firstColumn, lastRow, lastColumn);
            if (knownRanges.TryGetValue(key, out int existing)) {
                declarations[existing].Positions.Add(position);
                continue;
            }
            // An invalid rectangle still disqualifies overlapping valid merges. Its intersection
            // with the table is used only for conflict assessment, never exported as a clipped merge.
            var range = new IWorkTableMergeRange(firstRow + 1, firstColumn + 1,
                Math.Min(lastRow, rowCount - 1) + 1, Math.Min(lastColumn, columnCount - 1) + 1);
            knownRanges.Add(key, declarations.Count);
            declarations.Add(new MergeDeclaration(range, inBounds, position));
        }
        var conflicting = new HashSet<int>(IWorkMergeOverlapIndex.FindOverlapIndexes(
            declarations.Select(declaration => declaration.Range).ToArray(), columnCount, source.CancellationToken));
        foreach (int index in conflicting.OrderBy(index => index)) {
            foreach (int occurrence in declarations[index].Positions) {
                references.Declarations.Record(model, MergePairPath(occurrence) + "/2", 1,
                    IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            }
        }
        if (conflicting.Count > 0) MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
        // An unreadable range can conceal a conflict anywhere, so retain cells without applying merges.
        if (unknownRange) return Array.Empty<IWorkTableMergeRange>();
        return Array.AsReadOnly(declarations.Where((declaration, index) => declaration.InBounds && !conflicting.Contains(index))
            .Select(declaration => declaration.Range).ToArray());
    }

    private static string MergePairPath(int position) => "47/2/3[" + position.ToString(CultureInfo.InvariantCulture) + "]";

    private static void RecordMergeMessageFailure(IWorkArchiveRecord model, IWorkWireMessage message,
        int field, string path, IWorkSourceReferenceIssueCollector references) => references.Declarations.Record(model,
            path, message.FieldCount(field), message.FieldCount(field) != 1
                || message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)
                    ? IWorkSourceDeclarationIssueKind.RejectedMessageSet : IWorkSourceDeclarationIssueKind.MalformedMessage);

    private static void MarkMergeStorageUnsupported(IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_MERGE_UNSUPPORTED"
                && diagnostic.RecordIdentifier == model.Identifier)) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_MERGE_UNSUPPORTED",
            "An iWork table contains a malformed or unsupported merged range; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
    }

    private sealed class MergeDeclaration(IWorkTableMergeRange range, bool inBounds, int position) {
        internal IWorkTableMergeRange Range { get; } = range;
        internal bool InBounds { get; } = inBounds;
        internal List<int> Positions { get; } = new() { position };
    }
}
