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

            if (!source!.CellsFitWithinRange
                || !source.TryGetUsedBounds(out firstRow, out int firstColumn, out lastRow, out int lastColumn)
                || firstColumn < declaredFirstColumn) {
                source.Dispose();
                source = null;
                return false;
            }

            firstOrdinal = firstColumn - declaredFirstColumn;
            lastOrdinal = lastColumn - declaredFirstColumn;
            return true;
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
