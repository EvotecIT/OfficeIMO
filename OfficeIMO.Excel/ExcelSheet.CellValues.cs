using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Writes multiple cell values efficiently, using parallelization when beneficial.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        /// <param name="mode">Optional execution mode override.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <remarks>
        /// This is the canonical API for batch cell writes. Use this in place of the older
        /// <see cref="SetCellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, CancellationToken)"/>
        /// method, which will be removed in a future release.
        /// </remarks>
        public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (cells is null) {
                throw new ArgumentNullException(nameof(cells));
            }
            var list = cells as IList<(int Row, int Column, object Value)> ?? cells.ToList();
            if (list.Count == 0) return;

            // Single cell: trivially sequential
            if (list.Count == 1) {
                var single = list[0];
                CellValue(single.Row, single.Column, single.Value);
                return;
            }

            // Prepared buffers for parallel scenario
            var prepared = new (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[list.Count];
            var ssPlanner = new SharedStringPlanner();

            ExecuteWithPolicy(
                opName: "CellValues",
                itemCount: list.Count,
                overrideMode: mode,
                sequentialCore: () => {
                    // Sequential path - direct writes with NoLock
                    for (int i = 0; i < list.Count; i++) {
                        var (r, c, v) = list[i];
                        CellValueCore(r, c, v);
                    }
                },
                computeParallel: () => {
                    // Parallel compute phase - prepare values without DOM mutation
                    Parallel.For(0, list.Count, new ParallelOptions {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i => {
                        var (r, c, obj) = list[i];
                        var (val, type) = CoerceForCellNoDom(obj, ssPlanner);
                        prepared[i] = (r, c, val!, type!);
                    });
                },
                applySequential: () => {
                    // Apply phase - first fix shared strings, then write all values to DOM
                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list);
                },
                ct: ct
            );
        }

        /// <summary>
        /// Compute-only coercion for parallel scenarios. Does not mutate DOM.
        /// Uses <see cref="SharedStringPlanner"/> for string values.
        /// </summary>
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCellNoDom(object? value, SharedStringPlanner planner) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                s => {
                    var sanitized = planner.Note(s);
                    return new CellValue(sanitized);
                },
                dateTimeOffsetStrategy);
            return (cellValue, new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType));
        }

        /// <summary>
        /// Obsolete. Use <see cref="CellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, CancellationToken)"/> instead.
        /// </summary>
        [Obsolete("Use CellValues(...) instead.")]
        public void SetCellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default) {
            CellValues(cells, mode, ct);
        }

        private void ApplyPreparedCells(
            (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            IList<(int Row, int Column, object Value)> source) {
            var writer = new BatchCellWriter(this);

            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                var originalValue = source[i].Value;
                var cell = writer.GetOrCreateCell(p.Row, p.Col);
                cell.CellValue = p.Val;
                cell.DataType = p.Type;
                ApplyAutomaticCellFormatting(cell, originalValue, p.Type);
            }

            ClearHeaderCache();
        }

        private sealed class BatchCellWriter {
            private readonly ExcelSheet _sheet;
            private readonly SheetData _sheetData;
            private readonly Dictionary<int, BatchRowState> _rows;

            internal BatchCellWriter(ExcelSheet sheet) {
                _sheet = sheet;
                _sheetData = sheet.GetOrCreateSheetData();
                _rows = new Dictionary<int, BatchRowState>();

                foreach (var row in _sheetData.Elements<Row>()) {
                    if (row.RowIndex == null) {
                        continue;
                    }

                    _rows[(int)row.RowIndex.Value] = new BatchRowState(row);
                }
            }

            internal Cell GetOrCreateCell(int rowIndex, int columnIndex) {
                if (!_rows.TryGetValue(rowIndex, out BatchRowState? rowState)) {
                    var row = _sheet.GetOrCreateRowElement(_sheetData, rowIndex);
                    rowState = new BatchRowState(row);
                    _rows[rowIndex] = rowState;
                }

                return rowState.GetOrCreateCell(columnIndex, rowIndex);
            }

            private sealed class BatchRowState {
                private readonly Row _row;
                private readonly Dictionary<int, Cell> _cells;
                private Cell? _lastCell;
                private int _lastColumnIndex;

                internal BatchRowState(Row row) {
                    _row = row;
                    _cells = new Dictionary<int, Cell>();

                    foreach (var cell in row.Elements<Cell>()) {
                        var reference = cell.CellReference?.Value;
                        if (string.IsNullOrEmpty(reference)) {
                            continue;
                        }

                        int columnIndex = GetColumnIndex(reference!);
                        _cells[columnIndex] = cell;

                        if (columnIndex >= _lastColumnIndex) {
                            _lastColumnIndex = columnIndex;
                            _lastCell = cell;
                        }
                    }
                }

                internal Cell GetOrCreateCell(int columnIndex, int rowIndex) {
                    if (_cells.TryGetValue(columnIndex, out Cell? existing)) {
                        return existing;
                    }

                    string cellReference = GetColumnName(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);
                    var cell = new Cell { CellReference = cellReference };

                    if (_lastCell == null) {
                        var firstCell = _row.Elements<Cell>().FirstOrDefault();
                        if (firstCell != null) {
                            _row.InsertBefore(cell, firstCell);
                        } else {
                            _row.Append(cell);
                        }
                    } else if (columnIndex > _lastColumnIndex) {
                        _row.InsertAfter(cell, _lastCell);
                    } else {
                        Cell? insertAfter = null;
                        foreach (var existingCell in _row.Elements<Cell>()) {
                            var existingReference = existingCell.CellReference?.Value;
                            if (string.IsNullOrEmpty(existingReference)) {
                                continue;
                            }

                            int existingColumnIndex = GetColumnIndex(existingReference!);
                            if (existingColumnIndex > columnIndex) {
                                _row.InsertBefore(cell, existingCell);
                                _cells[columnIndex] = cell;
                                return cell;
                            }

                            insertAfter = existingCell;
                        }

                        if (insertAfter != null) {
                            _row.InsertAfter(cell, insertAfter);
                        } else {
                            _row.Append(cell);
                        }
                    }

                    _cells[columnIndex] = cell;
                    if (columnIndex >= _lastColumnIndex) {
                        _lastColumnIndex = columnIndex;
                        _lastCell = cell;
                    }

                    return cell;
                }
            }
        }
    }
}
