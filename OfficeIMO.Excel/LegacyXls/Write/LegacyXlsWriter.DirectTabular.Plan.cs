using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static bool TryCreateDirectTabularPlan(
            ExcelDirectTabularSource source,
            ExcelDateSystem dateSystem,
            CancellationToken cancellationToken,
            out DirectTabularPlan plan) {
            IExcelSheetTabularRowSource rows = source.Rows;
            int rowOffset = source.IncludeHeaders ? 1 : 0;
            int totalRows = checked(rows.RowCount + rowOffset);
            int columnCount = rows.ColumnCount;
            if (totalRows > 65_536 || columnCount > 256) {
                throw new NotSupportedException("Native XLS saving supports the BIFF8 worksheet limit of 65,536 rows and 256 columns.");
            }

            var firstColumns = new ushort[totalRows];
            for (int row = 0; row < firstColumns.Length; row++) {
                firstColumns[row] = ushort.MaxValue;
            }
            var lastColumns = new ushort[totalRows];
            var nonEmptyCellCounts = new ushort[totalRows];
            var sharedStrings = new LegacyXlsDirectSharedStringBuilder();
            int cellSlotCount = checked(totalRows * columnCount);
            int cachedCellSlotCount = cellSlotCount <= DirectMaximumCachedCellSlots
                ? cellSlotCount
                : 0;
            ulong[] rowFingerprints = cachedCellSlotCount == 0
                ? new ulong[totalRows]
                : Array.Empty<ulong>();
            byte[] cellKinds = cachedCellSlotCount == 0
                ? Array.Empty<byte>()
                : ArrayPool<byte>.Shared.Rent(cachedCellSlotCount);
            ulong[] cellPayloads = cachedCellSlotCount == 0
                ? Array.Empty<ulong>()
                : ArrayPool<ulong>.Shared.Rent(cachedCellSlotCount);
            bool returnCellCache = cachedCellSlotCount != 0;
            object?[]? flatValues = rows.TryGetFlatValues(out object?[] candidateValues, out int flatColumnCount)
                && flatColumnCount == columnCount
                && candidateValues.Length == checked(rows.RowCount * columnCount)
                    ? candidateValues
                    : null;
            int nonEmptyCellCount = 0;
            bool hasDates = false;
            uint? firstDataRow = null;
            uint? lastDataRow = null;

            try {
                if (source.IncludeHeaders) {
                    ulong rowFingerprint = DirectRowFingerprintOffset;
                    for (int column = 0; column < columnCount; column++) {
                        int cellSlot = column;
                        uint sharedStringIndex = sharedStrings.Add(rows.GetColumnName(column));
                        if (cachedCellSlotCount != 0) {
                            cellKinds[cellSlot] = (byte)ExcelDirectTabularValueKind.Text;
                            cellPayloads[cellSlot] = sharedStringIndex;
                        }
                        AddDirectCellFingerprint(
                            ref rowFingerprint,
                            ExcelDirectTabularValueKind.Text,
                            sharedStringIndex);
                        IncludeDirectCell(firstColumns, lastColumns, 0, checked((ushort)column));
                        nonEmptyCellCounts[0]++;
                        nonEmptyCellCount++;
                    }
                    if (cachedCellSlotCount == 0) rowFingerprints[0] = rowFingerprint;
                    if (columnCount != 0) {
                        firstDataRow = 0;
                        lastDataRow = 0;
                    }
                }

                bool canCancel = cancellationToken.CanBeCanceled;
                for (int row = 0; row < rows.RowCount; row++) {
                    if (canCancel && (row & 1023) == 0) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    int directRow = row + rowOffset;
                    ulong rowFingerprint = DirectRowFingerprintOffset;
                    object?[]? bufferedRow = flatValues == null
                        && rows.TryGetBufferedRow(row, out object?[]? candidateRow)
                        && candidateRow?.Length == columnCount
                            ? candidateRow
                            : null;
                    for (int column = 0; column < columnCount; column++) {
                        object? rawValue = flatValues != null
                            ? flatValues[checked((row * columnCount) + column)]
                            : bufferedRow != null
                                ? bufferedRow[column]
                                : rows.GetValue(row, column);
                        ExcelDirectTabularValue value = ExcelDirectTabularValue.Normalize(rawValue, dateSystem, source.PreserveMissingValues);
                        if (value.Kind == ExcelDirectTabularValueKind.Unsupported) {
                            plan = null!;
                            return false;
                        }

                        hasDates |= value.Kind == ExcelDirectTabularValueKind.DateTime;
                        int cellSlot = checked((directRow * columnCount) + column);
                        ulong payload = value.Kind switch {
                            ExcelDirectTabularValueKind.Text => sharedStrings.Add(value.Text ?? string.Empty),
                            ExcelDirectTabularValueKind.Boolean => value.Boolean ? 1UL : 0UL,
                            ExcelDirectTabularValueKind.Number or ExcelDirectTabularValueKind.DateTime => unchecked((ulong)BitConverter.DoubleToInt64Bits(value.Number)),
                            _ => 0UL
                        };
                        if (cachedCellSlotCount != 0) {
                            cellKinds[cellSlot] = (byte)value.Kind;
                            cellPayloads[cellSlot] = payload;
                        }
                        if (cachedCellSlotCount == 0) {
                            AddDirectCellFingerprint(ref rowFingerprint, value.Kind, payload);
                        }
                        if (value.Kind == ExcelDirectTabularValueKind.Empty) {
                            continue;
                        }

                        nonEmptyCellCount++;
                        nonEmptyCellCounts[directRow]++;
                        IncludeDirectCell(firstColumns, lastColumns, directRow, checked((ushort)column));
                        uint dataRow = checked((uint)directRow);
                        firstDataRow = firstDataRow.HasValue ? Math.Min(firstDataRow.Value, dataRow) : dataRow;
                        lastDataRow = lastDataRow.HasValue ? Math.Max(lastDataRow.Value, dataRow) : dataRow;
                    }
                    if (cachedCellSlotCount == 0) rowFingerprints[directRow] = rowFingerprint;
                }

                LegacyXlsDimensions dimensions = totalRows == 0 || columnCount == 0
                    ? default
                    : new LegacyXlsDimensions(0, checked((uint)totalRows), 0, checked((ushort)columnCount));
                if (totalRows != 0 && columnCount != 0) {
                    IncludeDirectCell(firstColumns, lastColumns, 0, 0);
                    IncludeDirectCell(
                        firstColumns,
                        lastColumns,
                        totalRows - 1,
                        checked((ushort)(columnCount - 1)));
                }

                int populatedRowCount = 0;
                for (int row = 0; row < firstColumns.Length; row++) {
                    if (firstColumns[row] != ushort.MaxValue) populatedRowCount++;
                }
                long estimatedCapacity = checked(
                    16_384L
                    + ((long)(nonEmptyCellCount + 2) * 18L)
                    + ((long)populatedRowCount * 22L)
                    + ((long)dimensions.RowBlockCount * 8L)
                    + sharedStrings.EstimatedSerializedBytes);
                int workbookStreamCapacity = estimatedCapacity >= int.MaxValue
                    ? int.MaxValue
                    : checked((int)estimatedCapacity);

                plan = new DirectTabularPlan(
                    totalRows,
                    columnCount,
                    cachedCellSlotCount,
                    cellKinds,
                    cellPayloads,
                    firstColumns,
                    lastColumns,
                    nonEmptyCellCounts,
                    rowFingerprints,
                    dimensions,
                    firstDataRow,
                    lastDataRow,
                    nonEmptyCellCount,
                    populatedRowCount,
                    workbookStreamCapacity,
                    sharedStrings.Build(),
                    rows,
                    flatValues,
                    rowOffset,
                    source.PreserveMissingValues,
                    dateSystem,
                    hasDates,
                    cancellationToken);
                returnCellCache = false;
                return true;
            } finally {
                if (returnCellCache) {
                    ClearAndReturnDirectCellCaches(cellKinds, cellPayloads);
                }
            }
        }

        internal static void ClearAndReturnDirectCellCaches(
            byte[] cellKinds,
            ulong[] cellPayloads,
            ArrayPool<byte>? cellKindPool = null,
            ArrayPool<ulong>? cellPayloadPool = null) {
            Array.Clear(cellKinds, 0, cellKinds.Length);
            Array.Clear(cellPayloads, 0, cellPayloads.Length);
            (cellKindPool ?? ArrayPool<byte>.Shared).Return(cellKinds);
            (cellPayloadPool ?? ArrayPool<ulong>.Shared).Return(cellPayloads);
        }

        private static void IncludeDirectCell(
            ushort[] firstColumns,
            ushort[] lastColumns,
            int row,
            ushort column) {
            if (firstColumns[row] == ushort.MaxValue) {
                firstColumns[row] = column;
                lastColumns[row] = column;
                return;
            }

            if (column < firstColumns[row]) firstColumns[row] = column;
            if (column > lastColumns[row]) lastColumns[row] = column;
        }

        private const ulong DirectRowFingerprintOffset = 14695981039346656037UL;
        private const ulong DirectRowFingerprintPrime = 1099511628211UL;

        private static void AddDirectCellFingerprint(
            ref ulong fingerprint,
            ExcelDirectTabularValueKind kind,
            ulong payload) {
            fingerprint = unchecked((fingerprint ^ (byte)kind) * DirectRowFingerprintPrime);
            fingerprint = unchecked((fingerprint ^ payload) * DirectRowFingerprintPrime);
        }

        private sealed class DirectTabularPlan : IDisposable {
            private int _disposed;

            internal DirectTabularPlan(
                int totalRows,
                int columnCount,
                int cellSlotCount,
                byte[] cellKinds,
                ulong[] cellPayloads,
                ushort[] firstColumns,
                ushort[] lastColumns,
                ushort[] nonEmptyCellCounts,
                ulong[] rowFingerprints,
                LegacyXlsDimensions dimensions,
                uint? firstDataRow,
                uint? lastDataRow,
                int nonEmptyCellCount,
                int populatedRowCount,
                int workbookStreamCapacity,
                LegacyXlsSharedStringTable sharedStrings,
                IExcelSheetTabularRowSource rows,
                object?[]? flatValues,
                int rowOffset,
                bool preserveMissingValues,
                ExcelDateSystem dateSystem,
                bool hasDates,
                CancellationToken cancellationToken) {
                TotalRows = totalRows;
                ColumnCount = columnCount;
                CellSlotCount = cellSlotCount;
                CellKinds = cellKinds;
                CellPayloads = cellPayloads;
                FirstColumns = firstColumns;
                LastColumns = lastColumns;
                NonEmptyCellCounts = nonEmptyCellCounts;
                RowFingerprints = rowFingerprints;
                Dimensions = dimensions;
                FirstDataRow = firstDataRow;
                LastDataRow = lastDataRow;
                NonEmptyCellCount = nonEmptyCellCount;
                PopulatedRowCount = populatedRowCount;
                WorkbookStreamCapacity = workbookStreamCapacity;
                SharedStrings = sharedStrings;
                Rows = rows;
                FlatValues = flatValues;
                RowOffset = rowOffset;
                PreserveMissingValues = preserveMissingValues;
                DateSystem = dateSystem;
                HasDates = hasDates;
                CancellationToken = cancellationToken;
            }

            internal int TotalRows { get; }
            internal int ColumnCount { get; }
            internal int CellSlotCount { get; }
            internal bool HasCachedCells => CellSlotCount != 0;
            internal byte[] CellKinds { get; }
            internal ulong[] CellPayloads { get; }
            internal ushort[] FirstColumns { get; }
            internal ushort[] LastColumns { get; }
            internal ushort[] NonEmptyCellCounts { get; }
            internal ulong[] RowFingerprints { get; }
            internal LegacyXlsDimensions Dimensions { get; }
            internal uint? FirstDataRow { get; }
            internal uint? LastDataRow { get; }
            internal int NonEmptyCellCount { get; }
            internal int PopulatedRowCount { get; }
            internal int WorkbookStreamCapacity { get; }
            internal LegacyXlsSharedStringTable SharedStrings { get; }
            internal IExcelSheetTabularRowSource Rows { get; }
            internal object?[]? FlatValues { get; }
            internal int RowOffset { get; }
            internal bool PreserveMissingValues { get; }
            internal ExcelDateSystem DateSystem { get; }
            internal bool HasDates { get; }
            internal ushort DateStyleIndex { get; set; }
            internal CancellationToken CancellationToken { get; }

            internal void ThrowIfCancellationRequested(int row) {
                if (CancellationToken.CanBeCanceled && (row & 1023) == 0) {
                    CancellationToken.ThrowIfCancellationRequested();
                }
            }

            public void Dispose() {
                if (CellSlotCount == 0 || Interlocked.Exchange(ref _disposed, 1) != 0) return;
                ClearAndReturnDirectCellCaches(CellKinds, CellPayloads);
            }
        }

        private static void PrepareUncachedDirectRow(
            DirectTabularPlan plan,
            int directRow,
            Span<byte> rowKinds,
            Span<ulong> rowPayloads) {
            plan.ThrowIfCancellationRequested(directRow);
            int firstColumn = int.MaxValue;
            int lastColumn = -1;
            int nonEmptyCellCount = 0;
            ulong rowFingerprint = DirectRowFingerprintOffset;
            int sourceRow = directRow - plan.RowOffset;
            object?[]? bufferedRow = sourceRow >= 0
                && plan.FlatValues == null
                && plan.Rows.TryGetBufferedRow(sourceRow, out object?[]? candidateRow)
                && candidateRow?.Length == plan.ColumnCount
                    ? candidateRow
                    : null;

            for (int column = 0; column < plan.ColumnCount; column++) {
                ExcelDirectTabularValue value;
                if (sourceRow < 0) {
                    value = ExcelDirectTabularValue.Normalize(plan.Rows.GetColumnName(column));
                } else {
                    object? rawValue = plan.FlatValues != null
                        ? plan.FlatValues[checked((sourceRow * plan.ColumnCount) + column)]
                        : bufferedRow != null
                            ? bufferedRow[column]
                            : plan.Rows.GetValue(sourceRow, column);
                    value = ExcelDirectTabularValue.Normalize(rawValue, plan.DateSystem, plan.PreserveMissingValues);
                }

                if (value.Kind == ExcelDirectTabularValueKind.Unsupported) {
                    throw new InvalidOperationException("The direct XLS tabular source changed after validation.");
                }

                rowKinds[column] = (byte)value.Kind;
                rowPayloads[column] = value.Kind switch {
                    ExcelDirectTabularValueKind.Text => GetValidatedDirectSharedStringIndex(
                        plan.SharedStrings,
                        value.Text ?? string.Empty),
                    ExcelDirectTabularValueKind.Boolean => value.Boolean ? 1UL : 0UL,
                    ExcelDirectTabularValueKind.Number or ExcelDirectTabularValueKind.DateTime => unchecked((ulong)BitConverter.DoubleToInt64Bits(value.Number)),
                    _ => 0UL
                };
                AddDirectCellFingerprint(ref rowFingerprint, value.Kind, rowPayloads[column]);
                if (value.Kind == ExcelDirectTabularValueKind.Empty) continue;

                nonEmptyCellCount++;
                if (firstColumn == int.MaxValue) firstColumn = column;
                lastColumn = column;
            }

            if (directRow == 0 && plan.ColumnCount != 0) {
                firstColumn = Math.Min(firstColumn, 0);
                lastColumn = Math.Max(lastColumn, 0);
            }
            if (directRow == plan.TotalRows - 1 && plan.ColumnCount != 0) {
                firstColumn = Math.Min(firstColumn, plan.ColumnCount - 1);
                lastColumn = Math.Max(lastColumn, plan.ColumnCount - 1);
            }

            ushort expectedFirstColumn = firstColumn == int.MaxValue
                ? ushort.MaxValue
                : checked((ushort)firstColumn);
            ushort expectedLastColumn = lastColumn < 0 ? (ushort)0 : checked((ushort)lastColumn);
            if (nonEmptyCellCount != plan.NonEmptyCellCounts[directRow]
                || expectedFirstColumn != plan.FirstColumns[directRow]
                || expectedLastColumn != plan.LastColumns[directRow]
                || rowFingerprint != plan.RowFingerprints[directRow]) {
                throw new InvalidOperationException("The direct XLS tabular source changed after validation.");
            }
        }

        private static uint GetValidatedDirectSharedStringIndex(
            LegacyXlsSharedStringTable sharedStrings,
            string text) {
            if (sharedStrings.TryGetDirectIndex(text, out uint index)) return index;
            throw new InvalidOperationException("The direct XLS tabular source changed after validation.");
        }

        private static void ValidateUncachedDirectRow(DirectTabularPlan plan, int directRow) {
            Span<byte> rowKinds = stackalloc byte[256];
            Span<ulong> rowPayloads = stackalloc ulong[256];
            PrepareUncachedDirectRow(plan, directRow, rowKinds, rowPayloads);
        }

    }
}
