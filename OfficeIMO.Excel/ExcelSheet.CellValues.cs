using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int DirectSequentialCellWriteLimit = 16;
        private const int DirectCellValuesLinearHeaderDuplicateCheckLimit = 32;
        private const int DirectCellValuesCompactReferenceCellThreshold = 4_096;
        private const int DirectCellValuesFlatSnapshotCellLimit = 500_000;

        /// <summary>
        /// Writes multiple cell values efficiently, using parallelization when beneficial.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        /// <param name="mode">Optional execution mode override.</param>
        /// <param name="ct">Cancellation token.</param>
        public void SetValues(IEnumerable<(int Row, int Column, object Value)> cells, ExcelExecutionMode? mode = null, CancellationToken ct = default) =>
            CellValuesCore(cells, mode, ct, preserveMissingValues: false);

        /// <summary>
        /// Writes a row-major batch on an empty sheet into the pending cell buffer, the same route single
        /// <see cref="CellValue(int, int, object)"/> calls take, so a sheet that cannot own the workbook's direct
        /// tabular model (any sheet after the first) still skips the cell DOM. Missing values stay blank.
        /// </summary>
        private bool TryBufferFreshSheetCellValues(IReadOnlyList<(int Row, int Column, object Value)> list, bool preserveMissingValues, CancellationToken ct) {
            if (_pendingCellValueDirectSaveBuffer != null || !CanAcceptPendingDirectCell()) return false;

            int previousRow = 0, previousColumn = 0;
            for (int i = 0; i < list.Count; i++) {
                var (row, column, value) = list[i];
                if (row <= 0 || column <= 0 || column > ExcelA1.MaxColumns
                    || row < previousRow || (row == previousRow && column <= previousColumn)) return false;
                previousRow = row;
                previousColumn = column;
                if (value is null || value is DBNull) {
                    if (!preserveMissingValues) return false;
                    continue;
                }
                if (value is DateTimeOffset) continue;
                if (!CanBufferDirectCell(ExcelBufferedCell.FromValue(value))) return false;
            }

            for (int i = 0; i < list.Count; i++) {
                ct.ThrowIfCancellationRequested();
                var (row, column, value) = list[i];
                if (value is null || value is DBNull) continue;
                if (!TrySetPendingDirectCell(row, column, ExcelBufferedCell.FromValue(value))) {
                    // Cells already written are rewritten with the same values by the DOM route.
                    MaterializePendingDirectCellValues(ct);
                    return false;
                }
            }

            ClearHeaderCache();
            return true;
        }

        /// <summary>
        /// Rejects writes into dynamic arrays before a batch is applied. The write index scans the worksheet and
        /// workbook parts, so it is built and read under the document write lock that concurrent batch writers apply under.
        /// </summary>
        /// <returns><see langword="true"/> when the sheet has dynamic array ranges.</returns>
        private bool EnsureCellValuesOutsideDynamicArrays(IReadOnlyList<(int Row, int Column, object Value)> list) =>
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                EnsureDynamicArrayWriteIndex();
                if (SpillOwnership.WriteRanges.Count == 0) return false;
                for (int i = 0; i < list.Count; i++) EnsureDynamicArrayCellWritable(list[i].Row, list[i].Column);
                return true;
            });

        /// <summary>
        /// Writes this sheet's pending buffer to the DOM before a batch writes its cells there. The batch looks for the
        /// buffer before it takes the document lock, so a batch on another thread can start one in between; left pending,
        /// it would stand in for the sheet at save with this batch's cells beside it.
        /// </summary>
        private void MaterializeConcurrentPendingDirectCellValues(CancellationToken ct) {
            if (_pendingCellValueDirectSaveBuffer != null
                && !_materializingPendingCellValueDirectSaveBuffer
                && !_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializePendingDirectCellValues(ct);
            }
        }

        private void CellValuesCore(IEnumerable<(int Row, int Column, object Value)> cells, ExcelExecutionMode? mode, CancellationToken ct, bool preserveMissingValues) {
            ArgumentNullException.ThrowIfNull(cells);
            var list = cells as IReadOnlyList<(int Row, int Column, object Value)> ?? cells.ToList();
            if (list.Count == 0) return;
            bool hasDynamicArrays = EnsureCellValuesOutsideDynamicArrays(list);
            DirectCellValuesSaveCandidate? appendSaveCandidate = null;

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                // Other sheets' pending cell buffers stay pending, as they do for single cell writes;
                // only this sheet's cells and the workbook's direct tabular model must reach the DOM first.
                if (_excelDocument.HasUnmaterializedDirectDataSetRows) {
                    _excelDocument.MaterializeDeferredDataSetImport();
                } else if (_pendingCellValueDirectSaveBuffer != null) {
                    if (HasCellValueWriteAccess) {
                        MaterializePendingDirectCellValues(ct);
                    } else {
                        Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => MaterializePendingDirectCellValues(ct));
                    }
                }
            }

            list = PrepareTableHeaderValueWrites(list, ct);

            // A batch may stand in for the whole sheet at save only if no other batch writes cells between the check
            // that the sheet is empty and the registration, so both run under the document lock other writers take.
            DirectCellValuesSaveCandidate? directSaveCandidate = null;
            bool registered = Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                if (TryCreateDirectCellValuesSaveCandidate(list, mode, out DirectCellValuesSaveCandidate? candidate)
                    && candidate != null
                    && CanRegisterDirectTabularSaveCandidate(1, 1, candidate.ColumnNames.Length)) {
                    directSaveCandidate = candidate;
                }

                if (directSaveCandidate != null
                    && RegisterDeferredDirectCellValuesSaveCandidateIfPossible(directSaveCandidate, preserveMissingValues)) {
                    return true;
                }

                if (appendSaveCandidate == null
                    && mode == ExcelExecutionMode.Parallel
                    && TryCreateDirectCellValuesAppendCandidate(list, out DirectCellValuesSaveCandidate? appendCandidate)) {
                    appendSaveCandidate = appendCandidate;
                }

                return appendSaveCandidate != null
                    && _excelDocument.CanDeferDirectCellValuesAppendCandidate
                    && RegisterDeferredDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate, preserveMissingValues);
            });
            if (registered) return;

            if (list.Count >= PendingDirectCellValueMinimumCellCount
                && !hasDynamicArrays
                && TryBufferFreshSheetCellValues(list, preserveMissingValues, ct)) {
                return;
            }

            // Single cell: trivially sequential
            if (list.Count == 1 && !preserveMissingValues) {
                var single = list[0];
                CellValue(single.Row, single.Column, single.Value);
                RegisterDirectCellValuesSaveCandidateIfPossible(directSaveCandidate);
                return;
            }

            if (!preserveMissingValues && list.Count > DirectSequentialCellWriteLimit && TryApplyPlainCellsByAppendingRows(list, ct)) {
                RegisterDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate ?? directSaveCandidate, preserveMissingValues);
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
                            if (preserveMissingValues && (v == null || v == DBNull.Value)) {
                                SetMissingTabularCellValue(GetWritableValueCell(r, c));
                                CompleteCellValueMutation(r, c);
                            } else {
                                CellValueCore(r, c, v);
                                if (preserveMissingValues) {
                                    // CellValueCore already admitted this value and synchronized any table header.
                                    ClearTabularReplacementValueMetadata(GetCell(r, c));
                                }
                            }
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
                    ApplyPreparedCells(prepared, list, preserveMissingValues);
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
                    MaterializeConcurrentPendingDirectCellValues(ct);
                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list, preserveMissingValues);
                },
                ct: ct
            );

            RegisterDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate ?? directSaveCandidate, preserveMissingValues);
        }

        /// <summary>
        /// Compute-only coercion for parallel scenarios. Does not mutate DOM.
        /// Uses <see cref="SharedStringPlanner"/> for string values.
        /// </summary>
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCellNoDom(object? value, SharedStringPlanner planner) {
            var dateTimeOffsetStrategy = _excelDocument.EffectiveDateTimeOffsetWriteStrategy;
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

    }
}
