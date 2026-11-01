using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using XCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Measures buffered cells through the same paths as worksheet cells, so the widths match
        // the DOM route and the cells can still reach a direct writer. Formats the buffer does
        // not reproduce are left to that route.
        private bool TryAutoFitPendingCells(IReadOnlyList<int>? columns, ExcelExecutionMode? mode, CancellationToken ct) {
            AutoFitMeasurementPlan? plan = null;
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => plan = TryBuildPendingCellAutoFitPlan(columns, ct));
            if (plan == null) {
                return false;
            }

            if (plan.Columns.Count > 0) {
                AutoFitColumnsInternal(plan, mode, ct);
            }

            return true;
        }

        private AutoFitMeasurementPlan? TryBuildPendingCellAutoFitPlan(IReadOnlyList<int>? columns, CancellationToken ct) {
            ExcelCellValueBuffer? buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null || buffer.CellCount < PendingDirectCellValueMinimumCellCount) {
                return null;
            }

            var targets = new Dictionary<int, int>();
            var columnsList = new List<int>();
            if (columns != null) {
                columnsList.AddRange(columns);
                for (int i = 0; i < columnsList.Count; i++) {
                    targets[columnsList[i]] = i;
                }
            }

            HashSet<string> sharedTexts = CollectPendingTextsThatFitSharedStrings();
            var measurements = new List<AutoFitMeasurement>();
            var uniqueMeasurements = new HashSet<AutoFitMeasurementKey>();
            var sharedStringMeasurements = new HashSet<(int TargetIndex, uint StyleIndex, int SharedStringId)>();
            var simpleTextMaxLengths = new Dictionary<(int TargetIndex, uint StyleIndex), int>();
            var textContext = CreateAutoFitTextContext();
            var scratchCells = new Dictionary<ExcelBufferedCellKind, Cell>();
            for (int rowIndex = 0; rowIndex < buffer.WrittenRowCount; rowIndex++) {
                ct.ThrowIfCancellationRequested();
                int end = buffer.GetEndCellIndex(rowIndex);
                for (int cellIndex = buffer.GetFirstCellIndex(rowIndex); cellIndex < end; cellIndex++) {
                    int column = buffer.GetColumn(cellIndex);
                    if (!targets.TryGetValue(column, out int targetIndex)) {
                        if (columns != null) {
                            continue;
                        }

                        targetIndex = columnsList.Count;
                        targets[column] = targetIndex;
                        columnsList.Add(column);
                    }

                    ExcelBufferedCell cell = buffer.GetCell(cellIndex);
                    if (cell.Kind == ExcelBufferedCellKind.Blank
                        || cell.Kind == ExcelBufferedCellKind.StyleOnly
                        || cell.Reference is DirectFormulaCellValue) {
                        continue;
                    }

                    uint style = GetEffectiveBufferedStyle(buffer, cellIndex);
                    if (cell.Reference is string text && (sharedTexts.Contains(text) || _excelDocument.TryGetExistingSharedStringIndex(text, out _, out _, out _))) {
                        if (!string.IsNullOrWhiteSpace(text) && uniqueMeasurements.Add(new AutoFitMeasurementKey(targetIndex, style, text, null))) {
                            measurements.Add(new AutoFitMeasurement(targetIndex, style, text, null));
                        }

                        continue;
                    }

                    if (!TryLoadAutoFitScratchCell(scratchCells, cell, style, out Cell? scratch)) {
                        return null;
                    }

                    AddAutoFitMeasurement(scratch!, targetIndex, textContext, uniqueMeasurements, sharedStringMeasurements, simpleTextMaxLengths, measurements);
                }
            }

            if (columns == null) {
                AddDeclaredColumns(targets, columnsList);
            }

            return new AutoFitMeasurementPlan(columnsList, measurements);
        }

        private void AddDeclaredColumns(Dictionary<int, int> targets, List<int> columnsList) {
            Columns? declared = WorksheetRoot.GetFirstChild<Columns>();
            if (declared == null) {
                return;
            }

            foreach (Column column in declared.Elements<Column>()) {
                uint max = column.Max?.Value ?? 0;
                for (uint index = column.Min?.Value ?? 0; index <= max; index++) {
                    if (!targets.ContainsKey((int)index)) {
                        targets[(int)index] = columnsList.Count;
                        columnsList.Add((int)index);
                    }
                }
            }
        }

        // Follows the DOM route, where text joins the shared table until it holds
        // CellValuePlainStringPromotionSharedStringCount entries and later text stays plain.
        private HashSet<string> CollectPendingTextsThatFitSharedStrings() {
            var texts = new HashSet<string>(StringComparer.Ordinal);
            int capacity = CellValuePlainStringPromotionSharedStringCount - _excelDocument.SharedStringCount;
            foreach (ExcelSheet sheet in _excelDocument.PendingDirectCellValueSheets) {
                foreach (string text in sheet._pendingCellValueDirectSaveBuffer!.EnumerateText()) {
                    if (texts.Count >= capacity) {
                        return texts;
                    }

                    texts.Add(text);
                }
            }

            return texts;
        }

        // One reusable cell per kind keeps measurement from allocating Open XML objects per cell.
        private bool TryLoadAutoFitScratchCell(Dictionary<ExcelBufferedCellKind, Cell> scratchCells, in ExcelBufferedCell cell, uint style, out Cell? scratch) {
            scratch = null;
            string raw;
            switch (cell.Kind) {
                case ExcelBufferedCellKind.Text:
                    raw = (string)cell.Reference!;
                    break;
                case ExcelBufferedCellKind.Number:
                    raw = InvariantNumberText.Get(cell.NumberValue);
                    break;
                case ExcelBufferedCellKind.Int32:
                    raw = InvariantNumberText.Get((int)cell.Bits);
                    break;
                case ExcelBufferedCellKind.Int64:
                    raw = InvariantNumberText.Get(cell.Bits);
                    break;
                case ExcelBufferedCellKind.Boolean:
                    raw = cell.Bits != 0L ? "1" : "0";
                    break;
                case ExcelBufferedCellKind.DateTime:
                    raw = InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(cell.DateTimeValue, _excelDocument.DateSystem));
                    break;
                default:
                    return false;
            }

            if (!scratchCells.TryGetValue(cell.Kind, out scratch)) {
                scratch = new Cell {
                    CellValue = new CellValue(raw),
                    DataType = cell.Kind == ExcelBufferedCellKind.Text ? XCellValues.String
                        : cell.Kind == ExcelBufferedCellKind.Boolean ? XCellValues.Boolean
                        : XCellValues.Number
                };
                scratchCells.Add(cell.Kind, scratch);
            } else {
                scratch.CellValue!.Text = raw;
            }

            if ((scratch.StyleIndex?.Value ?? 0U) != style) {
                scratch.StyleIndex = style;
            }

            return true;
        }
    }
}
