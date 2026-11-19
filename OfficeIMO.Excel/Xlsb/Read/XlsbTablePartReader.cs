using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Decodes Table parts (MS-XLSB 2.1.7.51) into canonical SpreadsheetML table definitions. Records outside the projected
    /// subset (sort state, XML maps, differential formats, cell styles, future records) are kept in the native part and
    /// mark the table as partially projected.
    /// </summary>
    internal static class XlsbTablePartReader {
        private const int BrtFRTBegin = 35;
        private const int BrtFRTEnd = 36;
        private const int BrtACBegin = 37;
        private const int BrtACEnd = 38;
        private const int BrtBeginAFilter = 161;
        private const int BrtEndAFilter = 162;
        private const int BrtBeginFilterColumn = 163;
        private const int BrtEndFilterColumn = 164;
        private const int BrtBeginFilters = 165;
        private const int BrtEndFilters = 166;
        private const int BrtFilter = 167;
        private const string TableRelationshipSuffix = "/relationships/table";
        private const string SingleCellsRelationshipSuffix = "/relationships/tableSingleCells";

        /// <summary>A table formula whose text is decoded once every table of the workbook is known.</summary>
        internal sealed class PendingFormula {
            internal PendingFormula(XlsbTable table, TableColumn column, bool totals, bool array, byte[] tokens, byte[] extra, int sheetIndex) {
                Table = table; Column = column; Totals = totals; Array = array; Tokens = tokens; Extra = extra; SheetIndex = sheetIndex;
            }
            internal XlsbTable Table { get; }
            internal TableColumn Column { get; }
            internal bool Totals { get; }
            internal bool Array { get; }
            internal byte[] Tokens { get; }
            /// <summary>The formula's rgcb, which holds PtgArray values.</summary>
            internal byte[] Extra { get; }
            internal int SheetIndex { get; }
        }

        /// <summary>True when <paramref name="relationship"/> targets a Table part.</summary>
        internal static bool IsTableRelationship(XlsbPackageRelationship relationship) =>
            !relationship.IsExternal
            && relationship.Type.StartsWith("http://schemas.openxmlformats.org/officeDocument/2006", StringComparison.Ordinal)
            && (relationship.Type.EndsWith(TableRelationshipSuffix, StringComparison.Ordinal)
                || relationship.Type.EndsWith(SingleCellsRelationshipSuffix, StringComparison.Ordinal));

        internal static XlsbTable Read(byte[] bytes, string relationshipId, string partName, int sheetIndex, XlsbImportOptions options,
            XlsbRecordReadBudget budget, List<PendingFormula> pending) {
            IReadOnlyList<XlsbRecord> records;
            using (var stream = new MemoryStream(bytes, writable: false)) {
                records = XlsbRecordReader.ReadAll(stream, options.MaxRecordBytes, budget, options.CancellationToken);
            }
            InvalidDataException Invalid(string detail) => new InvalidDataException($"The XLSB table part '{partName}' {detail}.");

            Table? table = null;
            XlsbCellRange? range = null;
            uint id = 0, headerRows = 0, totalsRows = 0;
            string? displayName = null;
            bool partial = false, sawColumns = false, sawStyle = false, ended = false;
            int depthAc = 0, depthFrt = 0;
            TableColumns? columns = null;
            TableColumn? column = null;
            var captions = new List<string>();
            var columnFormulas = new List<(TableColumn Column, bool Totals, bool Array, byte[] Tokens, byte[] Extra)>();
            XlsbAutoFilter? autoFilter = null;
            XlsbAutoFilterColumn? filterColumn = null;
            bool inFilters = false;
            foreach (XlsbRecord record in records) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (ended) throw Invalid("contains records after BrtEndList");
                // ACUID revision identifiers carry no table content.
                if (record.Type == BrtACBegin) { depthAc++; continue; }
                if (record.Type == BrtACEnd) { if (--depthAc < 0) throw Invalid("contains an unmatched BrtACEnd"); continue; }
                if (depthAc > 0) continue;
                if (record.Type == BrtFRTBegin) { depthFrt++; partial = true; continue; }
                if (record.Type == BrtFRTEnd) { if (--depthFrt < 0) throw Invalid("contains an unmatched BrtFRTEnd"); continue; }
                if (depthFrt > 0) continue;

                if (autoFilter != null && record.Type != BrtEndAFilter) {
                    switch (record.Type) {
                        case BrtBeginFilterColumn when filterColumn == null:
                            filterColumn = XlsbWorksheetAutoFilterReader.ReadColumn(record, autoFilter);
                            autoFilter.AddColumn(filterColumn);
                            break;
                        case BrtBeginFilters when filterColumn != null && !inFilters:
                            XlsbWorksheetAutoFilterReader.ReadBeginFilters(record, filterColumn);
                            inFilters = true;
                            break;
                        case BrtFilter when inFilters:
                            filterColumn!.AddValue(XlsbWorksheetAutoFilterReader.ReadValue(record));
                            break;
                        case BrtEndFilters when inFilters:
                            inFilters = false;
                            break;
                        case BrtEndFilterColumn when filterColumn != null && !inFilters && !filterColumn.InCustomFilters:
                            // A column the projection leaves out keeps the table partial, so an edit cannot drop it.
                            if (filterColumn.HasUnsupportedContent) {
                                autoFilter.HasUnsupportedContent = true;
                                partial = true;
                            }
                            filterColumn = null;
                            break;
                        case XlsbAutoFilterCriteria.BrtIconFilter when filterColumn != null && !inFilters:
                        case XlsbAutoFilterCriteria.BrtTop10Filter when filterColumn != null && !inFilters:
                        case XlsbAutoFilterCriteria.BrtDynamicFilter when filterColumn != null && !inFilters:
                        case XlsbAutoFilterCriteria.BrtBeginCustomFilters when filterColumn != null && !inFilters:
                        case XlsbAutoFilterCriteria.BrtEndCustomFilters when filterColumn != null && !inFilters:
                        case XlsbAutoFilterCriteria.BrtCustomFilter when filterColumn != null && !inFilters:
                            // The criteria the worksheet AutoFilter projects; colour filters stay native below.
                            XlsbWorksheetAutoFilterReader.ReadCriterion(record, filterColumn);
                            break;
                        default:
                            // Colour criteria and unknown records are kept natively.
                            if (filterColumn != null) filterColumn.HasUnsupportedContent = true;
                            autoFilter.HasUnsupportedContent = true;
                            partial = true;
                            break;
                    }
                    continue;
                }

                switch (record.Type) {
                    case XlsbWorksheetTableWriter.BrtBeginList: {
                        if (table != null) throw Invalid("contains more than one BrtBeginList");
                        var cursor = new XlsbBinaryCursor(record.Data);
                        int rfxOffset = 0;
                        range = XlsbWorksheetFeatureReader.ReadRfX(record.Data, ref rfxOffset, partName);
                        cursor.Skip(16);
                        uint listType = cursor.ReadUInt32();
                        id = cursor.ReadUInt32();
                        headerRows = cursor.ReadUInt32();
                        totalsRows = cursor.ReadUInt32();
                        uint flags = cursor.ReadUInt32();
                        bool hasDxf = false;
                        for (int index = 0; index < 6; index++) hasDxf |= cursor.ReadUInt32() != 0xFFFFFFFFU;
                        uint connection = cursor.ReadUInt32();
                        string? name = ReadNullable(cursor, options);
                        displayName = ReadNullable(cursor, options);
                        string? comment = ReadNullable(cursor, options);
                        bool hasCellStyles = false;
                        for (int index = 0; index < 3; index++) hasCellStyles |= ReadNullable(cursor, options) != null;
                        if (id < 1U || id > 0xFFFFFFFEU || headerRows > 1U || totalsRows > 1U || string.IsNullOrEmpty(displayName)
                            || range.LastRow - range.FirstRow + 1 <= headerRows + totalsRows) {
                            throw Invalid("declares an invalid table id, name, range or row counts");
                        }
                        partial |= listType != 0U || (flags & 0x02U) != 0 || hasDxf || connection != 0U || hasCellStyles;
                        table = new Table { Id = id };
                        if (name != null) table.Name = name;
                        table.DisplayName = displayName;
                        table.Reference = range.ToA1Reference();
                        table.HeaderRowCount = headerRows;
                        if (totalsRows != 0U) table.TotalsRowCount = totalsRows;
                        table.TotalsRowShown = (flags & 0x01U) != 0;
                        if ((flags & 0x04U) != 0) table.InsertRow = true;
                        if ((flags & 0x08U) != 0) table.InsertRowShift = true;
                        if ((flags & 0x10U) != 0) table.Published = true;
                        if (comment != null) table.Comment = comment;
                        break;
                    }
                    case BrtBeginAFilter when table != null && columns == null && autoFilter == null:
                        if (record.Data.Length != 16) throw Invalid("contains an invalid BrtBeginAFilter");
                        int filterOffset = 0;
                        autoFilter = new XlsbAutoFilter(XlsbWorksheetFeatureReader.ReadRfX(record.Data, ref filterOffset, partName));
                        break;
                    case BrtEndAFilter:
                        if (autoFilter == null || filterColumn != null || inFilters || table!.GetFirstChild<AutoFilter>() != null) {
                            throw Invalid("contains an unmatched AutoFilter end record");
                        }
                        table.Append(XlsbWorksheetAutoFilterProjector.Create(autoFilter));
                        autoFilter = null;
                        break;
                    case XlsbWorksheetTableWriter.BrtBeginListCols when table != null && !sawColumns:
                        sawColumns = true;
                        columns = new TableColumns();
                        break;
                    case XlsbWorksheetTableWriter.BrtBeginListCol when columns != null && column == null: {
                        var cursor = new XlsbBinaryCursor(record.Data);
                        uint columnId = cursor.ReadUInt32();
                        uint ilta = cursor.ReadUInt32();
                        bool hasDxf = false;
                        for (int index = 0; index < 3; index++) hasDxf |= cursor.ReadUInt32() != 0xFFFFFFFFU;
                        uint queryField = cursor.ReadUInt32();
                        string? uniqueName = ReadNullable(cursor, options);
                        string? caption = ReadNullable(cursor, options);
                        string? total = ReadNullable(cursor, options);
                        bool hasCellStyles = false;
                        for (int index = 0; index < 3; index++) hasCellStyles |= ReadNullable(cursor, options) != null;
                        if (columnId < 1U || string.IsNullOrEmpty(caption) || ilta > 9U) throw Invalid("declares an invalid table column");
                        partial |= hasDxf || queryField != 0U || uniqueName != null || hasCellStyles;
                        column = new TableColumn { Id = columnId, Name = caption };
                        if (ilta != 0U) column.TotalsRowFunction = XlsbWorksheetTableWriter.TotalsFunctions.First(entry => entry.Value == ilta).Function;
                        if (total != null) column.TotalsRowLabel = total;
                        captions.Add(caption!);
                        break;
                    }
                    case XlsbWorksheetTableWriter.BrtListCCFmla when column != null:
                    case XlsbWorksheetTableWriter.BrtListTrFmla when column != null: {
                        if (record.Data.Length < 9) throw Invalid("contains a truncated table formula");
                        bool array = (record.Data[0] & 0x02) != 0;
                        int cce = checked((int)BitConverter.ToUInt32(record.Data, 1));
                        if (cce < 0 || cce > record.Data.Length - 9) throw Invalid("contains a truncated table formula");
                        byte[] tokens = new byte[cce];
                        Buffer.BlockCopy(record.Data, 5, tokens, 0, cce);
                        int cb = checked((int)BitConverter.ToUInt32(record.Data, 5 + cce));
                        if (cb < 0 || cb > record.Data.Length - 9 - cce) throw Invalid("contains a truncated table formula");
                        byte[] extra = new byte[cb];
                        Buffer.BlockCopy(record.Data, 9 + cce, extra, 0, cb);
                        columnFormulas.Add((column, record.Type == XlsbWorksheetTableWriter.BrtListTrFmla, array, tokens, extra));
                        break;
                    }
                    case XlsbWorksheetTableWriter.BrtEndListCol when column != null:
                        columns!.Append(column);
                        column = null;
                        break;
                    case XlsbWorksheetTableWriter.BrtEndListCols when columns != null && column == null:
                        columns.Count = (uint)columns.ChildElements.Count;
                        table!.Append(columns);
                        columns = null;
                        break;
                    case XlsbWorksheetTableWriter.BrtTableStyleClient when sawColumns && columns == null && !sawStyle: {
                        sawStyle = true;
                        var cursor = new XlsbBinaryCursor(record.Data);
                        ushort flags = cursor.ReadUInt16();
                        string? style = ReadNullable(cursor, options);
                        var info = new TableStyleInfo();
                        if (style != null) info.Name = style;
                        info.ShowFirstColumn = (flags & 0x01) != 0;
                        info.ShowLastColumn = (flags & 0x02) != 0;
                        info.ShowRowStripes = (flags & 0x04) != 0;
                        info.ShowColumnStripes = (flags & 0x08) != 0;
                        table!.Append(info);
                        break;
                    }
                    case XlsbWorksheetTableWriter.BrtEndList when table != null && sawStyle:
                        ended = true;
                        break;
                    default:
                        // Sort state, XML column properties and other optional records stay in the native part.
                        if (table == null) throw Invalid("does not start with BrtBeginList");
                        partial = true;
                        break;
                }
            }
            if (!ended || depthAc != 0 || depthFrt != 0 || table == null || columns != null || column != null || autoFilter != null) {
                throw Invalid("is incomplete");
            }
            if (captions.Count != range!.LastColumn - range.FirstColumn + 1) throw Invalid("does not define one column per range column");

            var result = new XlsbTable(relationshipId, partName, table, id, displayName!, range, headerRows, totalsRows, captions, partial);
            foreach (var formula in columnFormulas) {
                pending.Add(new PendingFormula(result, formula.Column, formula.Totals, formula.Array, formula.Tokens, formula.Extra, sheetIndex));
            }
            return result;
        }

        /// <summary>Adds decodable calculated-column and totals formulas; an undecodable one leaves the table partial.</summary>
        internal static void ResolveFormulas(IEnumerable<PendingFormula> pending, XlsbWorkbook workbook) {
            foreach (PendingFormula formula in pending) {
                string? text;
                bool decoded;
                try {
                    decoded = XlsbFormulaTextReader.TryRead(formula.Tokens, out text, workbook, formula.SheetIndex, formula.Extra);
                } catch (NotSupportedException) {
                    decoded = false;
                    text = null;
                }
                if (!decoded || string.IsNullOrEmpty(text)) {
                    // The native part keeps the formula; the projection omits it and the table counts as partial.
                    formula.Table.HasUnsupportedContent = true;
                    continue;
                }
                TableFormulaType element = formula.Totals ? new TotalsRowFormula(text!) : new CalculatedColumnFormula(text!);
                if (formula.Array) element.Array = true;
                if (formula.Totals) formula.Column.Append(element);
                else formula.Column.InsertAt(element, 0);
            }
        }

        private static string? ReadNullable(XlsbBinaryCursor cursor, XlsbImportOptions options) {
            uint count = cursor.ReadUInt32();
            if (count == 0xFFFFFFFFU) return null;
            if (count > (uint)Math.Min(options.MaxStringCharacters, 32_767)) throw new InvalidDataException("A table string is too long.");
            return Encoding.Unicode.GetString(cursor.ReadBytes(checked((int)count * 2)));
        }
    }
}
