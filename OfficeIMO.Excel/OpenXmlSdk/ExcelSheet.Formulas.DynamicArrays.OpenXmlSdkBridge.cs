using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<
            WorksheetPart, DynamicSpillOwnership> DynamicSpillOwnerships =
            new System.Runtime.CompilerServices.ConditionalWeakTable<
                WorksheetPart, DynamicSpillOwnership>();

        private sealed class DynamicSpillOwnership {
            internal readonly Dictionary<long, DynamicSpillCacheSnapshot> Cells =
                new Dictionary<long, DynamicSpillCacheSnapshot>();
            internal bool WriteIndexInitialized;
            internal readonly List<(int Top, int Left, int Bottom, int Right)> WriteRanges =
                new List<(int Top, int Left, int Bottom, int Right)>();
            internal bool OriginalFingerprintAttempted;
            internal byte[]? OriginalFingerprint;
            internal bool OriginalScanAttempted;
            internal readonly HashSet<long> WrittenOwners = new HashSet<long>();
        }

        private DynamicSpillOwnership? _spillOwnership;

        private DynamicSpillOwnership SpillOwnership => _spillOwnership ??= DynamicSpillOwnerships.GetOrCreateValue(_worksheetPart);

        private Dictionary<long, DynamicSpillCacheSnapshot> DynamicSpillCacheSnapshots =>
            SpillOwnership.Cells;

        private sealed class DynamicSpillCacheSnapshot {
            internal Cell? SourceCell;
            internal CellValues? Type;
            internal string? Value;
            internal uint? ValueMetaIndex;
        }

        private static DynamicSpillCacheSnapshot SnapshotDynamicSpillCell(Cell cell) => new DynamicSpillCacheSnapshot {
            SourceCell = cell,
            Type = cell.DataType?.Value,
            Value = ExcelWorksheetScan.FindValue(cell)?.Text,
            ValueMetaIndex = cell.ValueMetaIndex?.Value
        };

        private static bool MatchesDynamicSpillSnapshot(Cell cell, DynamicSpillCacheSnapshot snapshot) =>
            cell.DataType?.Value == snapshot.Type
            && string.Equals(ExcelWorksheetScan.FindValue(cell)?.Text, snapshot.Value, StringComparison.Ordinal)
            && cell.ValueMetaIndex?.Value == snapshot.ValueMetaIndex;

        private void InvalidateDynamicArrayWriteIndex() => SpillOwnership.WriteIndexInitialized = false;

        private void EnsureDynamicArrayWriteIndex() {
            DynamicSpillOwnership ownership = SpillOwnership;
            if (ownership.WriteIndexInitialized) return;
            ownership.WriteRanges.Clear();
            var part = _excelDocument.WorkbookPartRoot.CellMetadataPart;
            if (part != null) {
                if (!part.IsRootElementLoaded) ValidateInCellImageMetadataPart(part, "Cell metadata");
                Metadata? metadata = part.Metadata;
                foreach (Cell cell in ExcelWorksheetScan.FormulaCells(WorksheetRoot)) {
                    if (ExcelWorksheetScan.FindFormula(cell)?.FormulaType?.Value != CellFormulaValues.Array
                        || cell.CellMetaIndex == null
                        || CreateFormulaArrayInfo(cell, metadata)?.IsDynamic != true
                        || !TryFixedArrayBounds(ExcelWorksheetScan.FindFormula(cell)!.Reference?.Value ?? "",
                            out int top, out int left, out int bottom, out int right)) continue;
                    ownership.WriteRanges.Add((top, left, bottom, right));
                }
            }
            ownership.WriteIndexInitialized = true;
        }

        private void EnsureDynamicArrayCellWritable(int row, int column) {
            EnsureDynamicArrayWriteIndex();
            List<(int Top, int Left, int Bottom, int Right)> ranges = SpillOwnership.WriteRanges;
            for (int i = 0; i < ranges.Count; i++) {
                var range = ranges[i];
                if (row >= range.Top && row <= range.Bottom && column >= range.Left && column <= range.Right)
                    throw new InvalidOperationException($"Cell '{ExcelA1.CellReference(row, column)}' belongs to a dynamic array; clear its array formula first.");
            }
        }

        internal Cell GetWritableValueCell(int row, int column) {
            EnsureCellPayloadWritable(row, column);
            return GetCell(row, column);
        }

        private void EnsureDynamicArrayRangeWritable(int top, int left, int bottom, int right) {
            EnsureDynamicArrayWriteIndex();
            if (SpillOwnership.WriteRanges.Any(range =>
                RangesOverlapInclusive((top, left, bottom, right), range)))
                throw new InvalidOperationException("The range overlaps a dynamic array; clear its array formula first.");
        }

        private void WriteDynamicArrayFormulaCache(DynamicArrayPlan plan) {
            Cell anchor = (Cell)plan.Owner.Cell.Element!;
            SpillOwnership.WrittenOwners.Add(DynamicCellKey(plan.Owner.Top, plan.Owner.Left));
            string? previousReference = ExcelWorksheetScan.FindFormula(anchor)?.Reference?.Value;
            if (plan.ErrorSubtype != 0) {
                ExcelWorksheetScan.FindFormula(anchor)!.Reference = anchor.CellReference!.Value!;
                TryWriteRichFormulaError(anchor, "#SPILL!", plan);
            } else {
                FormulaArrayValue array = plan.Array!;
                ExcelWorksheetScan.FindFormula(anchor)!.Reference = ExcelA1.CellReference(plan.Top, plan.Left) + ":"
                    + ExcelA1.CellReference(plan.Bottom, plan.Right);
                for (int row = plan.Top; row <= plan.Bottom; row++)
                    for (int column = plan.Left; column <= plan.Right; column++) {
                        Cell cell = row == plan.Top && column == plan.Left ? anchor : GetCell(row, column);
                        cell.InlineString = null;
                        cell.ValueMetaIndex = null;
                        SetFormulaCachedValue(cell, array.Values[(row - plan.Top) * array.Columns + column - plan.Left]);
                        if (cell != anchor)
                            DynamicSpillCacheSnapshots[DynamicCellKey(row, column)] = SnapshotDynamicSpillCell(cell);
                    }
            }
            FixedArraySheetIndex index = GetFixedArraySheetIndex();
            for (int row = plan.Owner.Top; row <= plan.Owner.Bottom; row++)
                for (int column = plan.Owner.Left; column <= plan.Owner.Right; column++) {
                    if (row == plan.Owner.Top && column == plan.Owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    if (!index.Cells.TryGetValue(DynamicCellKey(row, column), out ExcelFormulaCell handle)) continue;
                    Cell cell = (Cell)handle.Element!;
                    long key = DynamicCellKey(row, column);
                    if (!IsOwnedDynamicSpillCachedCell(cell, key)) continue;
                    cell.CellValue = null;
                    cell.InlineString = null;
                    cell.DataType = null;
                    cell.ValueMetaIndex = null;
                    DynamicSpillCacheSnapshots.Remove(key);
                    if (cell.StyleIndex == null && !cell.HasChildren) cell.Remove();
                }
            if (!string.Equals(previousReference, ExcelWorksheetScan.FindFormula(anchor)?.Reference?.Value, StringComparison.Ordinal))
                InvalidateDynamicArrayWriteIndex();
        }

        private bool IsDynamicSpillObstacle(FixedArrayOwner owner, ExcelFormulaCell handle, int row, int column,
            FormulaArgumentValue expected) {
            Cell cell = (Cell)handle.Element!;
            if (row >= owner.Top && row <= owner.Bottom && column >= owner.Left && column <= owner.Right)
                return !IsOwnedDynamicSpillCachedCell(cell, DynamicCellKey(row, column), expected);
            return ExcelWorksheetScan.FindValue(cell) != null || cell.InlineString != null || cell.ValueMetaIndex != null
                || ExcelWorksheetScan.FindFormula(cell) != null || cell.CellMetaIndex != null;
        }

        private bool IsOwnedDynamicSpillCachedCell(ExcelFormulaCell handle, long key, FormulaArgumentValue? expected = null) {
            Cell cell = (Cell)handle.Element!;
            if (ExcelWorksheetScan.FindFormula(cell) != null || cell.CellMetaIndex != null || cell.InlineString != null) return false;
            if (DynamicSpillCacheSnapshots.TryGetValue(key, out DynamicSpillCacheSnapshot? snapshot))
                return MatchesDynamicSpillSnapshot(cell, snapshot);
            if (ExcelWorksheetScan.FindValue(cell) == null && cell.ValueMetaIndex == null) return true;
            if (expected == null) return false;
            FormulaArgumentValue value = expected.Value;
            if (value.IsError)
                return cell.DataType?.Value == CellValues.Error
                    && ResolveRichValueError(cell, ExcelWorksheetScan.FindValue(cell)?.Text) == value.ErrorCode;
            if (cell.ValueMetaIndex != null) return false;
            if (value.IsBoolean)
                return cell.DataType?.Value == CellValues.Boolean
                    && ExcelWorksheetScan.FindValue(cell)?.Text == (value.Number == 0 ? "0" : "1");
            if (value.Number.HasValue)
                return (cell.DataType == null
                        || cell.DataType.Value == CellValues.Number)
                    && OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(ExcelWorksheetScan.FindValue(cell)?.Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double numeric)
                    && numeric == value.Number.Value;
            return value.Text != null && string.Equals(GetCellText(cell), value.Text, StringComparison.Ordinal);
        }

        /// <summary>The ranges of the sheet's tables.</summary>
        private IEnumerable<string> EnumerateTableRanges() {
            foreach (TableDefinitionPart part in _worksheetPart.TableDefinitionParts)
                yield return part.Table?.Reference?.Value ?? string.Empty;
        }
    }
}
