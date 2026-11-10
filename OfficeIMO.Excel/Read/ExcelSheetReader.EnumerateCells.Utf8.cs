namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>
        /// Enumerates cells from the indexed UTF-8 worksheet source used by the tabular readers.
        /// A worksheet that source declines keeps the XML reader route.
        /// </summary>
        private bool TryCreateEnumerateCellsUtf8Source(
            CancellationToken ct,
            out ExcelUtf8RangeRowSource? source,
            out int firstRow,
            out int lastRow,
            out int firstOrdinal,
            out int lastOrdinal,
            out int declaredFirstColumn) {
            source = null;
            firstRow = lastRow = firstOrdinal = lastOrdinal = declaredFirstColumn = 0;
            if (_opt.CellValueConverter != null
                || !ReferenceEquals(_opt.Culture, System.Globalization.CultureInfo.InvariantCulture)) {
                return false;
            }

            try {
                if (!ExcelUtf8RangeRowSource.TryCreateForUsedRange(this, ct, out source, out declaredFirstColumn)) return false;
            } catch (Exception ex) when (ex is NotSupportedException or InvalidDataException or System.Xml.XmlException or FormatException) {
                // Shared-formula followers, invalid references and malformed XML keep the XML
                // route, which reports them the way cell enumeration always has.
                source = null;
                return false;
            }

            if (!source!.TryGetUsedBounds(out firstRow, out int firstColumn, out lastRow, out int lastColumn)) {
                source.Dispose();
                source = null;
                return false;
            }

            if (!source.CellsFitWithinRange || firstColumn < declaredFirstColumn) {
                // A worksheet without a dimension is proposed from its first row; a later row that starts further
                // left falls outside it. The scan still found the real bounds, so index exactly those.
                source.Dispose();
                source = null;
                if (!TryCreateEnumerateCellsUtf8Source(firstRow, lastRow, firstColumn, lastColumn, ct, out source)) return false;
                declaredFirstColumn = firstColumn;
            }

            firstOrdinal = firstColumn - declaredFirstColumn;
            lastOrdinal = lastColumn - declaredFirstColumn;
            return true;
        }

        private bool TryCreateEnumerateCellsUtf8Source(
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn,
            CancellationToken ct,
            out ExcelUtf8RangeRowSource? source) {
            source = null;
            try {
                if (!ExcelUtf8RangeRowSource.TryCreate(this, firstRow, lastRow, firstColumn, lastColumn - firstColumn + 1, ct, out source)) return false;
            } catch (Exception ex) when (ex is NotSupportedException or InvalidDataException or System.Xml.XmlException or FormatException) {
                source = null;
                return false;
            }
            if (source!.CellsFitWithinRange) return true;
            source.Dispose();
            source = null;
            return false;
        }

        private static IEnumerable<ExcelCellValueInfo> EnumerateCellsUtf8(
            ExcelUtf8RangeRowSource source,
            int firstRow,
            int lastRow,
            int firstOrdinal,
            int lastOrdinal,
            int declaredFirstColumn,
            CancellationToken ct) {
            using (source) {
                bool canCancel = ct.CanBeCanceled;
                for (int row = firstRow; row <= lastRow; row++) {
                    if (canCancel) ct.ThrowIfCancellationRequested();
                    if (!source.SelectRow(row, ct, ct)) continue;

                    for (int ordinal = firstOrdinal; ordinal <= lastOrdinal; ordinal++) {
                        if (!source.IsCellPresent(ordinal)) continue;
                        source.ReadValue(ordinal, XmlDataReaderTargetKind.None,
                            out _, out _, out _, out _, out _, out _, out object? value);
                        if (value != null) {
                            yield return new ExcelCellValueInfo(row, ordinal + declaredFirstColumn, value);
                        }
                    }
                }
            }
        }
    }
}
