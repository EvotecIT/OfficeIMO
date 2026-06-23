using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal sealed class BiffSharedFormulaImportState {
        private readonly LegacyXlsWorksheet _sheet;
        private readonly IReadOnlyList<BiffExternSheetReference> _externSheets;
        private readonly IReadOnlyList<string> _sheetNames;
        private readonly IReadOnlyList<string?> _definedNames;
        private readonly List<LegacyXlsImportDiagnostic> _diagnostics;
        private readonly LegacyXlsImportOptions _options;
        private readonly Dictionary<SharedFormulaAnchor, SharedFormulaDefinition> _definitions = new();
        private readonly List<PendingSharedFormulaCell> _pendingCells = new();
        private PendingSharedFormulaCell? _lastSharedFormulaCell;

        internal BiffSharedFormulaImportState(
            LegacyXlsWorksheet sheet,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options) {
            _sheet = sheet;
            _externSheets = externSheets;
            _sheetNames = sheetNames;
            _definedNames = definedNames;
            _diagnostics = diagnostics;
            _options = options;
        }

        internal static bool TryReadFormulaReference(byte[] formulaPayload, int parsedFormulaOffset, out BiffSharedFormulaReference reference) {
            reference = default;
            if (parsedFormulaOffset + 7 > formulaPayload.Length) {
                return false;
            }

            ushort expressionLength = BiffRecordReader.ReadUInt16(formulaPayload, parsedFormulaOffset);
            if (expressionLength != 5 || formulaPayload[parsedFormulaOffset + 2] != 0x01) {
                return false;
            }

            reference = new BiffSharedFormulaReference(
                BiffRecordReader.ReadUInt16(formulaPayload, parsedFormulaOffset + 3),
                BiffRecordReader.ReadUInt16(formulaPayload, parsedFormulaOffset + 5));
            return true;
        }

        internal void RegisterFormulaCell(int row, int column, BiffSharedFormulaReference? reference, int recordOffset) {
            if (reference == null) {
                _lastSharedFormulaCell = null;
                return;
            }

            var pendingCell = new PendingSharedFormulaCell(row, column, reference.Value.Anchor, recordOffset);
            _lastSharedFormulaCell = pendingCell;
            if (_definitions.TryGetValue(reference.Value.Anchor, out SharedFormulaDefinition definition)) {
                ResolveCell(pendingCell, definition);
                return;
            }

            _pendingCells.Add(pendingCell);
        }

        internal bool TryReadDefinition(byte[] payload, int recordOffset) {
            PendingSharedFormulaCell? lastSharedFormulaCell = _lastSharedFormulaCell;
            _lastSharedFormulaCell = null;
            if (payload.Length < 10 || lastSharedFormulaCell == null) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, 0);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, 2);
            ushort firstColumn = payload[4];
            ushort lastColumn = payload[5];
            ushort expressionLength = BiffRecordReader.ReadUInt16(payload, 8);
            if (lastRow < firstRow
                || lastColumn < firstColumn
                || expressionLength == 0
                || 10 + expressionLength > payload.Length) {
                return false;
            }

            PendingSharedFormulaCell anchorCell = lastSharedFormulaCell.Value;
            if (!ContainsCell(firstRow, lastRow, firstColumn, lastColumn, anchorCell.Row - 1, anchorCell.Column - 1)) {
                return false;
            }

            byte[] formulaPayload = new byte[checked(2 + expressionLength)];
            Buffer.BlockCopy(payload, 8, formulaPayload, 0, formulaPayload.Length);
            var definition = new SharedFormulaDefinition(
                anchorCell.Anchor,
                firstRow,
                lastRow,
                firstColumn,
                lastColumn,
                formulaPayload,
                recordOffset);
            _definitions[anchorCell.Anchor] = definition;
            ResolvePendingCells(definition);
            return true;
        }

        internal void AddUnresolvedDiagnostics() {
            if (!_options.ReportUnsupportedRecords) {
                return;
            }

            foreach (PendingSharedFormulaCell cell in _pendingCells) {
                if (_definitions.ContainsKey(cell.Anchor)) {
                    continue;
                }

                _diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    "XLS-BIFF-FORMULA-SHARED-UNRESOLVED",
                    $"Shared formula for {A1.ColumnIndexToLetters(cell.Column)}{cell.Row} was not found; the cached result was imported without formula text.",
                    sheetName: _sheet.Name,
                    recordOffset: cell.RecordOffset,
                    recordType: (ushort)BiffRecordType.Formula,
                    detailCode: "FormulaSharedFormulaUnresolved"));
            }
        }

        private void ResolvePendingCells(SharedFormulaDefinition definition) {
            for (int i = _pendingCells.Count - 1; i >= 0; i--) {
                PendingSharedFormulaCell cell = _pendingCells[i];
                if (cell.Anchor.Equals(definition.Anchor) && definition.Contains(cell.Row - 1, cell.Column - 1)) {
                    ResolveCell(cell, definition);
                    _pendingCells.RemoveAt(i);
                }
            }
        }

        private void ResolveCell(PendingSharedFormulaCell cell, SharedFormulaDefinition definition) {
            if (!definition.Contains(cell.Row - 1, cell.Column - 1)) {
                return;
            }

            if (BiffFormulaTextReader.TryRead(
                definition.FormulaPayload,
                0,
                cell.Row - 1,
                cell.Column - 1,
                _externSheets,
                _sheetNames,
                _definedNames,
                out string? formulaText,
                out BiffFormulaReadFailure? failure)) {
                if (!string.IsNullOrWhiteSpace(formulaText)) {
                    _sheet.TryReplaceFormulaText(cell.Row, cell.Column, formulaText!);
                }

                return;
            }

            if (!_options.ReportUnsupportedRecords) {
                return;
            }

            string failureDescription = failure == null ? "Unsupported shared formula tokens" : failure.Description;
            _diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Info,
                "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED",
                $"{failureDescription} Shared formula at {A1.ColumnIndexToLetters(cell.Column)}{cell.Row} was imported from its cached result.",
                sheetName: _sheet.Name,
                recordOffset: definition.RecordOffset,
                recordType: (ushort)BiffRecordType.ShrFmla,
                detailCode: failure?.DetailCode));
        }

        private static bool ContainsCell(ushort firstRow, ushort lastRow, ushort firstColumn, ushort lastColumn, int row, int column) {
            return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
        }

        private readonly struct SharedFormulaDefinition {
            internal SharedFormulaDefinition(
                SharedFormulaAnchor anchor,
                ushort firstRow,
                ushort lastRow,
                ushort firstColumn,
                ushort lastColumn,
                byte[] formulaPayload,
                int recordOffset) {
                Anchor = anchor;
                FirstRow = firstRow;
                LastRow = lastRow;
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
                FormulaPayload = formulaPayload;
                RecordOffset = recordOffset;
            }

            internal SharedFormulaAnchor Anchor { get; }

            internal ushort FirstRow { get; }

            internal ushort LastRow { get; }

            internal ushort FirstColumn { get; }

            internal ushort LastColumn { get; }

            internal byte[] FormulaPayload { get; }

            internal int RecordOffset { get; }

            internal bool Contains(int row, int column) {
                return ContainsCell(FirstRow, LastRow, FirstColumn, LastColumn, row, column);
            }
        }

        private readonly struct PendingSharedFormulaCell {
            internal PendingSharedFormulaCell(int row, int column, SharedFormulaAnchor anchor, int recordOffset) {
                Row = row;
                Column = column;
                Anchor = anchor;
                RecordOffset = recordOffset;
            }

            internal int Row { get; }

            internal int Column { get; }

            internal SharedFormulaAnchor Anchor { get; }

            internal int RecordOffset { get; }
        }
    }

    internal readonly struct BiffSharedFormulaReference {
        internal BiffSharedFormulaReference(ushort anchorRow, ushort anchorColumn) {
            Anchor = new SharedFormulaAnchor(anchorRow, anchorColumn);
        }

        internal SharedFormulaAnchor Anchor { get; }
    }

    internal readonly struct SharedFormulaAnchor : IEquatable<SharedFormulaAnchor> {
        internal SharedFormulaAnchor(ushort row, ushort column) {
            Row = row;
            Column = column;
        }

        internal ushort Row { get; }

        internal ushort Column { get; }

        public bool Equals(SharedFormulaAnchor other) {
            return Row == other.Row && Column == other.Column;
        }

        public override bool Equals(object? obj) {
            return obj is SharedFormulaAnchor other && Equals(other);
        }

        public override int GetHashCode() {
            unchecked {
                return (Row * 397) ^ Column;
            }
        }
    }
}
