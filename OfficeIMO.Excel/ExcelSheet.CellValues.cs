using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int DirectSequentialCellWriteLimit = 16;
        private const int DirectCellValuesLinearHeaderDuplicateCheckLimit = 32;
        private const int DirectCellValuesFlatSnapshotCellLimit = 500_000;

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
            var list = cells as IReadOnlyList<(int Row, int Column, object Value)> ?? cells.ToList();
            if (list.Count == 0) return;
            DirectCellValuesSaveCandidate? appendSaveCandidate = null;

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            DirectCellValuesSaveCandidate? directSaveCandidate = null;
            if (TryCreateDirectCellValuesSaveCandidate(list, mode, out DirectCellValuesSaveCandidate? candidate)
                && candidate != null
                && CanRegisterDirectTabularSaveCandidate(1, 1, candidate.ColumnNames.Length)) {
                directSaveCandidate = candidate;
            }

            if (directSaveCandidate != null
                && RegisterDeferredDirectCellValuesSaveCandidateIfPossible(directSaveCandidate)) {
                return;
            }

            if (appendSaveCandidate == null
                && mode == ExecutionMode.Parallel
                && TryCreateDirectCellValuesAppendCandidate(list, out DirectCellValuesSaveCandidate? appendCandidate)) {
                appendSaveCandidate = appendCandidate;
            }

            if (appendSaveCandidate != null
                && _excelDocument.CanDeferDirectCellValuesAppendCandidate
                && RegisterDeferredDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate)) {
                return;
            }

            // Single cell: trivially sequential
            if (list.Count == 1) {
                var single = list[0];
                CellValue(single.Row, single.Column, single.Value);
                RegisterDirectCellValuesSaveCandidateIfPossible(directSaveCandidate);
                return;
            }

            if (list.Count > DirectSequentialCellWriteLimit && TryApplyPlainCellsByAppendingRows(list, ct)) {
                RegisterDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate ?? directSaveCandidate);
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
                    if (list.Count <= DirectSequentialCellWriteLimit) {
                        for (int i = 0; i < list.Count; i++) {
                            ct.ThrowIfCancellationRequested();
                            var (r, c, v) = list[i];
                            CellValueCore(r, c, v);
                        }

                        return;
                    }

                    // Sequential path - keep the fast prepared/apply writer so row-major
                    // batches can append rows instead of falling back to GetCell per cell.
                    for (int i = 0; i < list.Count; i++) {
                        ct.ThrowIfCancellationRequested();
                        var (r, c, v) = list[i];
                        var (val, type) = CoerceForCellNoDom(v, ssPlanner);
                        prepared[i] = (r, c, val!, type!);
                    }

                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list);
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

            RegisterDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate ?? directSaveCandidate);
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
                dateTimeOffsetStrategy,
                _excelDocument.DateSystem);
            return (cellValue, GetCachedDataTableCellType(cellType));
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
