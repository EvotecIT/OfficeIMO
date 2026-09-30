using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static bool HasUnsupportedTableScalarEncoding(IWorkWireMessage message) =>
        new[] { 6, 7, 9, 10, 11, 16, 17 }.Any(field => message.FieldCount(field) > 1)
        || message.HasUnexpectedWireKind(6, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(7, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(9, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(10, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(11, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(16, IWorkWireKind.Fixed64)
        || message.HasUnexpectedWireKind(17, IWorkWireKind.Fixed64);

    private static IReadOnlyList<IWorkTableMergeRange> ReadMergedRanges(IWorkWireMessage table,
        int rowCount, int columnCount, int maximumRanges, int maximumFormulaNodes,
        IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        IWorkWireMessage? mergeOwner = IWorkObjectIndex.TryGetMessage(table, 47, out bool malformedOwner);
        if (malformedOwner) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (mergeOwner == null) return Array.Empty<IWorkTableMergeRange>();
        byte[]? formulaStoreBytes = mergeOwner.GetBytes(2);
        int pairCount;
        try {
            pairCount = formulaStoreBytes == null
                || mergeOwner.FieldCount(2) != 1
                || mergeOwner.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                    ? -1
                    : mergeOwner.CountNestedFields(formulaStoreBytes, 3);
        } catch (InvalidDataException) {
            pairCount = -1;
        }
        if (pairCount > maximumRanges) {
            throw new InvalidDataException($"iWork table merged-range count exceeds the configured limit of {maximumRanges} in object {model.Identifier}.");
        }
        if (pairCount < 0) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        IWorkWireMessage formulaStore;
        try {
            formulaStore = mergeOwner.ParseNestedMessage(formulaStoreBytes!);
        } catch (InvalidDataException) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        IReadOnlyList<IWorkWireMessage> pairs = IWorkObjectIndex.TryGetMessages(formulaStore, 3, out bool malformedPairs);
        if (malformedPairs) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (pairs.Count != pairCount) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        var result = new List<IWorkTableMergeRange>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (IWorkWireMessage pair in pairs) {
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(pair, 2, out bool malformedFormula);
            if (malformedFormula || formula == null
                || !IWorkFormulaReader.TryReadAbsoluteRange(formula, maximumFormulaNodes,
                    out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                || firstRow >= rowCount || lastRow >= rowCount
                || firstColumn >= columnCount || lastColumn >= columnCount) {
                MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            if (firstRow == lastRow && firstColumn == lastColumn) continue;
            string key = firstRow.ToString(CultureInfo.InvariantCulture) + ":"
                + firstColumn.ToString(CultureInfo.InvariantCulture) + ":"
                + lastRow.ToString(CultureInfo.InvariantCulture) + ":"
                + lastColumn.ToString(CultureInfo.InvariantCulture);
            if (!seen.Add(key)) continue;
            result.Add(new IWorkTableMergeRange(firstRow + 1, firstColumn + 1, lastRow + 1, lastColumn + 1));
        }
        if (IWorkMergeRangeValidator.HasOverlaps(result, columnCount)) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
        }
        return Array.AsReadOnly(result.ToArray());
    }

    private static void MarkMergeStorageUnsupported(IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_MERGE_UNSUPPORTED"
                && diagnostic.RecordIdentifier == model.Identifier)) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_TABLE_MERGE_UNSUPPORTED",
            "An iWork table contains a malformed or unsupported merged range; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
    }

    private static int CheckedSubDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the table dimensions.");
        }
        return (int)resolved;
    }

    private static double? ValidDimension(double? value) => value.HasValue && IsFinite(value.Value) && value.Value > 0
        ? value
        : null;

    private static bool HasInvalidDeclaredDimension(IWorkWireMessage message, int field,
        double? value) => message.FieldCount(field) > 1
        || message.HasField(field)
            && (message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed64)
                || !value.HasValue || !IsFinite(value.Value) || value.Value <= 0);

    private static int CheckedDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the configured limit of {maximum}.");
        }
        return (int)resolved;
    }

}
