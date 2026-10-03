namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    private static string BindTableReference(IWorkWireMessage node,
        int? firstColumn, int? lastColumn, int? firstRow, int? lastRow,
        bool firstColumnAbsolute, bool lastColumnAbsolute, bool firstRowAbsolute, bool lastRowAbsolute,
        int maximumCharacters, IReadOnlyDictionary<Guid, IWorkFormulaTableBinding>? tableQualifiers,
        IWorkFormulaTableBinding? owningTable, ref bool complete, ref bool boundedBodyRanges) {
        string address = ReferenceAddress(firstColumn, lastColumn, firstRow, lastRow,
            firstColumnAbsolute, lastColumnAbsolute, firstRowAbsolute, lastRowAbsolute);
        if (address == "#REF!") { complete = false; return address; }
        bool wholeColumns = !firstRow.HasValue, wholeRows = !firstColumn.HasValue;
        IWorkFormulaTableBinding? target;
        if (node.HasField(28)) {
            Guid? identifier = ReadReferencedTableIdentifier(node);
            if (!identifier.HasValue || tableQualifiers == null
                || !tableQualifiers.TryGetValue(identifier.Value, out target))
                return PreserveUnresolvedTableReference(node, address, ref complete);
        } else {
            if (!wholeColumns && !wholeRows || owningTable == null) return address;
            target = owningTable;
        }
        if (wholeColumns || wholeRows) {
            IWorkTable table = target.Table;
            if (!table.BodyMetadataIsComplete
                || wholeColumns && (!firstColumn.HasValue || lastColumn >= table.ColumnCount)
                || wholeRows && (!firstRow.HasValue || lastRow >= table.RowCount)) {
                complete = false; return "#REF!";
            }
            if (target.BoundBodyRanges) {
                if (wholeColumns) {
                    firstRow = table.HeaderRowCount; lastRow = table.RowCount - table.FooterRowCount - 1;
                    firstRowAbsolute = lastRowAbsolute = true;
                } else {
                    firstColumn = table.HeaderColumnCount; lastColumn = table.ColumnCount - 1;
                    firstColumnAbsolute = lastColumnAbsolute = true;
                }
                address = ReferenceAddress(firstColumn, lastColumn, firstRow, lastRow,
                    firstColumnAbsolute, lastColumnAbsolute, firstRowAbsolute, lastRowAbsolute);
                if (address == "#REF!") { complete = false; return address; }
                boundedBodyRanges = true;
            }
        }
        return Bound(target.Qualifier + address, maximumCharacters, ref complete);
    }

    private static string ReferenceAddress(int? firstColumn, int? lastColumn, int? firstRow, int? lastRow,
        bool firstColumnAbsolute, bool lastColumnAbsolute, bool firstRowAbsolute, bool lastRowAbsolute) {
        if (firstColumn.HasValue != lastColumn.HasValue || firstRow.HasValue != lastRow.HasValue
            || !firstColumn.HasValue && !firstRow.HasValue || lastColumn < firstColumn || lastRow < firstRow)
            return "#REF!";
        string first = CellAddress(firstColumn, firstRow, firstColumnAbsolute, firstRowAbsolute);
        string last = CellAddress(lastColumn, lastRow, lastColumnAbsolute, lastRowAbsolute);
        if (first == "#REF!" || last == "#REF!") return "#REF!";
        // Even a single whole row/column needs a colon in the coordinate source representation.
        return first == last && firstColumn.HasValue && firstRow.HasValue ? first : first + ":" + last;
    }
}
