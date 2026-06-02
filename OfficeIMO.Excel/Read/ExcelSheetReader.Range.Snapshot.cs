using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range-based read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private void FillSequentialBuffers(
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            int startRow,
            object?[]? headerValues,
            object?[][] rowValues,
            CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            int visitedCells = 0;
            int inferredRowIndex = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = GetSequentialRowIndex(row, ref inferredRowIndex);
                if (rowIndex < r1) continue;
                if (rowIndex > r2) continue;

                int rr = rowIndex - r1 - startRow;
                bool isHeaderRow = headerValues != null && rowIndex == r1;
                if (!isHeaderRow && (uint)rr >= (uint)rowValues.Length) {
                    continue;
                }

                object?[]? values = null;
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++visitedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    int cc = cIndex - c1;
                    if ((uint)cc >= (uint)cols) continue;

                    if (TryConvertCell(cell, out object? value)) {
                        if (isHeaderRow) {
                            headerValues![cc] = value;
                        } else {
                            values ??= rowValues[rr] ??= new object?[cols];
                            values[cc] = value;
                        }
                    }
                }
            }
        }

        private static Type[] InferDataTableColumnTypes(object?[][] rowValues, int cols) {
            var types = new Type[cols];
            for (int c = 0; c < cols; c++) {
                Type? inferred = null;
                for (int r = 0; r < rowValues.Length; r++) {
                    object?[]? row = rowValues[r];
                    if (row == null) {
                        continue;
                    }

                    inferred = MergeDataTableColumnType(inferred, row[c]);
                    if (inferred == typeof(object)) {
                        break;
                    }
                }

                types[c] = inferred ?? typeof(object);
            }

            return types;
        }

        private static Type[] InferDataTableColumnTypes(object?[,] values, int startRow, int rows, int cols) {
            var types = new Type[cols];
            for (int c = 0; c < cols; c++) {
                Type? inferred = null;
                for (int r = startRow; r < rows; r++) {
                    inferred = MergeDataTableColumnType(inferred, values[r, c]);
                    if (inferred == typeof(object)) {
                        break;
                    }
                }

                types[c] = inferred ?? typeof(object);
            }

            return types;
        }

        private static Type[] InferDataTableColumnTypesFromRaw(List<CellRaw> raw, int r1, int c1, int cols, int startRow) {
            var types = new Type[cols];
            Type?[] inferred = new Type?[cols];
            for (int i = 0; i < raw.Count; i++) {
                var cell = raw[i];
                int rr = cell.Row - r1 - startRow;
                int cc = cell.Col - c1;
                if (rr < 0 || (uint)cc >= (uint)cols || inferred[cc] == typeof(object)) {
                    continue;
                }

                inferred[cc] = MergeDataTableColumnType(inferred[cc], cell.TypedValue);
            }

            for (int c = 0; c < cols; c++) {
                types[c] = inferred[c] ?? typeof(object);
            }

            return types;
        }

        private static Type? MergeDataTableColumnType(Type? current, object? value) {
            if (value == null || value == DBNull.Value) {
                return current;
            }

            Type next = value.GetType();
            if (current == null || current == next) {
                return next;
            }

            return typeof(object);
        }

        private List<CellRaw> SnapshotAndConvertRangeCells(
            int r1,
            int c1,
            int r2,
            int c2,
            string operationName,
            OfficeIMO.Excel.ExecutionMode? mode,
            CancellationToken ct,
            int workload) {
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            var raw = new List<CellRaw>(capacity: GetSnapshotCapacity(workload));
            SnapshotCellsInto(raw, r1, c1, r2, c2, ct, out bool needsSharedStrings, out bool needsStyles);
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) decided = policy.Decide(operationName, raw.Count);

            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel && raw.Count > 0) {
                PrepareCachesForParallelConversion(needsSharedStrings, needsStyles);
                var po = new ParallelOptions {
                    CancellationToken = ct,
                    MaxDegreeOfParallelism = policy.MaxDegreeOfParallelism ?? -1
                };
                Parallel.For(0, raw.Count, po, i => raw[i] = ConvertRaw(raw[i]));
            } else {
                bool canCancel = ct.CanBeCanceled;
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                for (int i = 0; i < raw.Count; i++) {
                    if (canCancel && (i & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    raw[i] = ConvertRaw(raw[i]);
                }
            }

            return raw;
        }

        private void PrepareCachesForParallelConversion(bool needsSharedStrings, bool needsStyles) {
            if (needsSharedStrings) {
                _sst.EnsureLoaded();
            }

            if (needsStyles) {
                _ = Styles;
            }
        }

        private static int GetSnapshotCapacity(int workload) {
            if (workload <= 0) {
                return 0;
            }

            if (workload <= DenseSnapshotCapacityLimit) {
                return workload;
            }

            return Math.Max(1024, workload / 4);
        }

        private void SnapshotCellsInto(List<CellRaw> buffer, int r1, int c1, int r2, int c2, CancellationToken ct, out bool needsSharedStrings, out bool needsStyles) {
            needsSharedStrings = false;
            needsStyles = false;
            bool canCancel = ct.CanBeCanceled;
            int visitedCells = 0;
            int inferredRowIndex = 0;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = GetSequentialRowIndex(row, ref inferredRowIndex);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++visitedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    var raw = SnapshotCell(cell, rIndex, cIndex);

                    if (raw.RawText != null || raw.InlineText != null || raw.FormulaText != null || CellHasExplicitBlank(cell) || _opt.FillBlanksInRanges) {
                        buffer.Add(raw);
                        if (!needsSharedStrings && raw.TypeHint == CellValues.SharedString) {
                            needsSharedStrings = true;
                        }

                        if (!needsStyles && _opt.TreatDatesUsingNumberFormat && raw.StyleIndex is not null) {
                            needsStyles = true;
                        }
                    }
                }
            }
        }

        private void FillRangeSequential(object?[,] result, int r1, int c1, int r2, int c2, CancellationToken ct) {
            int height = result.GetLength(0);
            int width = result.GetLength(1);
            bool canCancel = ct.CanBeCanceled;
            int visitedCells = 0;
            int inferredRowIndex = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = GetSequentialRowIndex(row, ref inferredRowIndex);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                int rr = rIndex - r1;
                if ((uint)rr >= (uint)height) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++visitedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    int cc = cIndex - c1;
                    if ((uint)cc >= (uint)width) continue;

                    if (TryConvertCell(cell, out object? value))
                        result[rr, cc] = value;
                }
            }
        }

        private static int GetSequentialRowIndex(Row row, ref int inferredRowIndex) {
            if (row.RowIndex != null) {
                inferredRowIndex = checked((int)row.RowIndex.Value);
            } else {
                inferredRowIndex++;
            }

            return inferredRowIndex;
        }

        private bool TryFillRangeXmlFast(object?[,] result, int r1, int c1, int r2, int c2, CancellationToken ct) {
            if (!CanAttemptXmlFastReader()) {
                return false;
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                int width = result.GetLength(1);
                int height = result.GetLength(0);
                var seenRows = CreateCompletedRowTracker(height);
                object?[]? rowBuffer8 = width == 8 ? new object?[8] : null;
                bool orderedRows = true;
                int orderedRowsSeen = 0;
                if (canCancel) {
                    while (reader.Read()) {
                        ct.ThrowIfCancellationRequested();

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                            continue;
                        }

                        int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                        if (rowIndex <= 0) {
                            rowIndex = nextRowIndex;
                        }

                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == height : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        ReadXmlRowIntoRange(reader, result, rowIndex, r1, c1, c2, width, rowBuffer8, ct);
                        if (orderedRows && rowIndex == r1 + orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == height) {
                                break;
                            }

                            continue;
                        }

                        if (orderedRows) {
                            for (int row = 0; row < orderedRowsSeen; row++) {
                                seenRows.MarkSeen(row);
                            }

                            orderedRows = false;
                        }

                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                } else {
                    while (reader.Read()) {
                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                            continue;
                        }

                        int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                        if (rowIndex <= 0) {
                            rowIndex = nextRowIndex;
                        }

                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == height : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        ReadXmlRowIntoRange(reader, result, rowIndex, r1, c1, c2, width, rowBuffer8, CancellationToken.None);
                        if (orderedRows && rowIndex == r1 + orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == height) {
                                break;
                            }

                            continue;
                        }

                        if (orderedRows) {
                            for (int row = 0; row < orderedRowsSeen; row++) {
                                seenRows.MarkSeen(row);
                            }

                            orderedRows = false;
                        }

                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                }

                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }

    }
}
