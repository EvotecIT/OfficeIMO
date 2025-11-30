using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
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
            var wrapFlags = new bool[list.Count];
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
                        if (type?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString && val?.Text is string raw) {
                            if (raw.Contains("\n") || raw.Contains("\r"))
                                wrapFlags[i] = true;
                        }
                        prepared[i] = (r, c, val!, type!);
                    });
                },
                applySequential: () => {
                    // Apply phase - first fix shared strings, then write all values to DOM
                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    for (int i = 0; i < prepared.Length; i++) {
                        var p = prepared[i];
                        var cell = GetCell(p.Row, p.Col);
                        cell.CellValue = p.Val;
                        cell.DataType = p.Type;
                        if (wrapFlags[i]) {
                            ApplyWrapText(cell);
                        }
                    }
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
    }
}
