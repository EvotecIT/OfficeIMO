namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    private static string BindTableReference(IWorkWireMessage node,
        int? firstColumn, int? lastColumn, int? firstRow, int? lastRow,
        bool firstColumnAbsolute, bool lastColumnAbsolute, bool firstRowAbsolute, bool lastRowAbsolute,
        int maximumCharacters, IReadOnlyDictionary<Guid, IWorkFormulaTableBinding>? tableQualifiers,
        IWorkFormulaTableBinding? owningTable, ref bool complete, ref bool boundedBodyRanges,
        out IntegerReference? integerReference) {
        integerReference = null;
        string address = ReferenceAddress(firstColumn, lastColumn, firstRow, lastRow,
            firstColumnAbsolute, lastColumnAbsolute, firstRowAbsolute, lastRowAbsolute,
            collapseSingleCell: node.GetUnsigned(1) == 36);
        if (address == "#REF!") { complete = false; return address; }
        bool wholeColumns = !firstRow.HasValue, wholeRows = !firstColumn.HasValue;
        IWorkFormulaTableBinding? target;
        if (node.HasField(28)) {
            Guid? identifier = ReadReferencedTableIdentifier(node);
            if (!identifier.HasValue || tableQualifiers == null
                || !tableQualifiers.TryGetValue(identifier.Value, out target))
                return PreserveUnresolvedTableReference(node, address, ref complete);
        } else {
            target = owningTable;
            if (!wholeColumns && !wholeRows) {
                integerReference = new IntegerReference(target?.Table, firstColumn!.Value, lastColumn!.Value,
                    firstRow!.Value, lastRow!.Value, singleCellSyntax: node.GetUnsigned(1) == 36);
                return address;
            }
            if (target == null) return address;
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
                    firstColumnAbsolute, lastColumnAbsolute, firstRowAbsolute, lastRowAbsolute,
                    collapseSingleCell: node.GetUnsigned(1) == 36);
                if (address == "#REF!") { complete = false; return address; }
                boundedBodyRanges = true;
            }
        }
        // Whole-axis operands retain their existing rendering contract, but are not
        // part of the bounded integer-reference qualification.
        if (!wholeColumns && !wholeRows)
            integerReference = new IntegerReference(target.Table, firstColumn!.Value, lastColumn!.Value,
                firstRow!.Value, lastRow!.Value, singleCellSyntax: node.GetUnsigned(1) == 36);
        return Bound(target.Qualifier + address, maximumCharacters, ref complete);
    }

    private static string ReferenceAddress(int? firstColumn, int? lastColumn, int? firstRow, int? lastRow,
        bool firstColumnAbsolute, bool lastColumnAbsolute, bool firstRowAbsolute, bool lastRowAbsolute,
        bool collapseSingleCell) {
        if (firstColumn.HasValue != lastColumn.HasValue || firstRow.HasValue != lastRow.HasValue
            || !firstColumn.HasValue && !firstRow.HasValue || lastColumn < firstColumn || lastRow < firstRow)
            return "#REF!";
        string first = CellAddress(firstColumn, firstRow, firstColumnAbsolute, firstRowAbsolute);
        string last = CellAddress(lastColumn, lastRow, lastColumnAbsolute, lastRowAbsolute);
        if (first == "#REF!" || last == "#REF!") return "#REF!";
        // A native range keeps its colon even when both endpoints coincide:
        // integer aggregates distinguish a blank direct cell from a blank range.
        return collapseSingleCell && first == last && firstColumn.HasValue && firstRow.HasValue
            ? first : first + ":" + last;
    }
}
