using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Encodes worksheet tables as Table parts (MS-XLSB 2.1.7.51) and the worksheet's LISTPARTS collection
    /// (2.4.103, 2.4.707, 2.4.457).
    /// </summary>
    internal static class XlsbWorksheetTableWriter {
        internal const string ContentType = "application/vnd.ms-excel.table";
        internal const string RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
        internal const int BrtBeginList = 343;
        internal const int BrtEndList = 344;
        internal const int BrtBeginListCols = 345;
        internal const int BrtEndListCols = 346;
        internal const int BrtBeginListCol = 347;
        internal const int BrtEndListCol = 348;
        internal const int BrtListCCFmla = 351;
        internal const int BrtListTrFmla = 352;
        internal const int BrtTableStyleClient = 513;
        internal const int BrtBeginListParts = 660;
        internal const int BrtListPart = 661;
        internal const int BrtEndListParts = 662;
        private const uint NoDxf = 0xFFFFFFFFU;
        private const int MaximumNameLength = 255;
        private const int MaximumTotalLabelLength = 8189;

        // ListTotalRowFunction (MS-XLSB 2.5.88) values for CT_TableColumn totalsRowFunction.
        internal static readonly (TotalsRowFunctionValues Function, uint Value)[] TotalsFunctions = {
            (TotalsRowFunctionValues.None, 0U),
            (TotalsRowFunctionValues.Average, 1U),
            (TotalsRowFunctionValues.Count, 2U),
            (TotalsRowFunctionValues.CountNumbers, 3U),
            (TotalsRowFunctionValues.Maximum, 4U),
            (TotalsRowFunctionValues.Minimum, 5U),
            (TotalsRowFunctionValues.Sum, 6U),
            (TotalsRowFunctionValues.StandardDeviation, 7U),
            (TotalsRowFunctionValues.Variance, 8U),
            (TotalsRowFunctionValues.Custom, 9U),
        };

        private static readonly string[] TableAttributes = {
            "id", "name", "displayName", "ref", "tableType", "headerRowCount", "totalsRowCount", "totalsRowShown",
            "insertRow", "insertRowShift", "published", "comment",
        };

        private static readonly string[] ColumnAttributes = { "id", "name", "totalsRowFunction", "totalsRowLabel" };

        private static readonly string[] StyleAttributes = { "name", "showFirstColumn", "showLastColumn", "showRowStripes", "showColumnStripes" };

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) =>
            records.TableParts.AddRange(Create(context));

        /// <summary>Validates and encodes the worksheet's tables, adding one Table part per table to <see cref="XlsbWorksheetFeatureContext.Parts"/>.</summary>
        internal static IReadOnlyList<XlsbGeneratedRecord> Create(XlsbWorksheetFeatureContext context) {
            string sheetName = context.Sheet.Name;
            WorksheetPart worksheetPart = context.Sheet.WorksheetPart;
            TableParts[] collections = context.Worksheet.Elements<TableParts>().ToArray();
            TableDefinitionPart[] definitions = worksheetPart.TableDefinitionParts.ToArray();
            if (collections.Length == 0 && definitions.Length == 0) return Array.Empty<XlsbGeneratedRecord>();
            if (collections.Length != 1
                || collections[0].ChildElements.Any(child => child is not TablePart)
                || XlsbOpenXmlAttributeValidator.FindUnsupported(collections[0], new[] { "count" }) is OpenXmlAttribute) {
                throw new NotSupportedException($"Native XLSB generation requires one plain tableParts collection on worksheet '{sheetName}'.");
            }

            TablePart[] references = collections[0].Elements<TablePart>().ToArray();
            var listed = new HashSet<TableDefinitionPart>();
            var tables = new List<(TableDefinitionPart Part, Table Table)>(references.Length);
            foreach (TablePart reference in references) {
                string? id = reference.Id?.Value;
                if (string.IsNullOrEmpty(id) || reference.HasChildren
                    || XlsbOpenXmlAttributeValidator.FindUnsupported(reference, Array.Empty<string>(), IsRelationshipId) is OpenXmlAttribute
                    || !worksheetPart.TryGetPartById(id!, out OpenXmlPart? part)
                    || part is not TableDefinitionPart definition || !listed.Add(definition)) {
                    throw new NotSupportedException($"Native XLSB generation found a table reference that does not name one table part on worksheet '{sheetName}'.");
                }
                if (definition.Parts.Any() || definition.ExternalRelationships.Any() || definition.HyperlinkRelationships.Any()) {
                    throw new NotSupportedException($"Native XLSB generation does not support query tables or other parts related to a table on worksheet '{sheetName}'.");
                }
                Table table = definition.Table
                    ?? throw new NotSupportedException($"Native XLSB generation found an empty table part on worksheet '{sheetName}'.");
                tables.Add((definition, table));
            }
            if (definitions.Any(definition => !listed.Contains(definition))) {
                throw new NotSupportedException($"Native XLSB generation found a table part that the worksheet does not list on worksheet '{sheetName}'.");
            }
            if (tables.Count == 0) return Array.Empty<XlsbGeneratedRecord>();

            // Encode every table before adding parts so a refusal leaves the package plan unchanged.
            var encoded = tables.Select(item => CreateTablePart(context, item.Table)).ToArray();
            string[] sourceIds = references.Select(reference => reference.Id!.Value!).ToArray();
            XlsbWorksheetPackageParts parts = context.Parts
                ?? throw new NotSupportedException($"Native XLSB rewriting cannot add table parts to worksheet '{sheetName}'; save as .xlsx to retain the tables.");
            var result = new List<XlsbGeneratedRecord>(encoded.Length + 2) {
                new XlsbGeneratedRecord(BrtBeginListParts, UInt32Payload(checked((uint)encoded.Length)))
            };
            for (int index = 0; index < encoded.Length; index++) {
                byte[] content = encoded[index];
                string partName = parts.Names.Next("xl/tables/table", ".bin");
                parts.AddPart(new XlsbGeneratedPart(partName, ContentType, content));
                string relationshipId = parts.AddRelationship(RelationshipType, partName, sourceIds[index]);
                using var payload = new MemoryStream(4 + relationshipId.Length * 2);
                XlsbPayloadWriter.WriteNullableWideString(payload, relationshipId);
                result.Add(new XlsbGeneratedRecord(BrtListPart, payload.ToArray()));
            }
            result.Add(new XlsbGeneratedRecord(BrtEndListParts, Array.Empty<byte>()));
            return result;
        }

        /// <summary>
        /// Re-encodes the Table parts of a loaded worksheet whose definitions changed since load, keeping their part names and
        /// relationships. Unchanged tables keep their native bytes.
        /// </summary>
        internal static void Rewrite(XlsbWorksheetFeatureContext context, Model.XlsbSheet source, IDictionary<string, byte[]> replacements,
            ISet<string> generatedParts) {
            if (!Projection.XlsbWorksheetTableProjector.TryPair(context.Sheet, source, out var pairs)) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add or remove tables on worksheet '{context.Sheet.Name}'; save as .xlsx to retain that change.");
            }
            foreach (var (part, table) in pairs) {
                Table current = part.Table
                    ?? throw new NotSupportedException($"Native XLSB rewriting found an empty table part on worksheet '{context.Sheet.Name}'.");
                if (string.Equals(current.OuterXml, table.Definition.OuterXml, StringComparison.Ordinal)) continue;
                if (table.HasUnsupportedContent) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot change table '{table.DisplayName}' on worksheet '{context.Sheet.Name}' because its native part holds content outside the supported table subset; save as .xlsx to retain that change.");
                }
                if (current.Id?.Value != table.Id) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot change the id of table '{table.DisplayName}' on worksheet '{context.Sheet.Name}'; formulas refer to tables by id.");
                }
                // Retained PtgList tokens elsewhere address columns by position; keep every column id where it was.
                uint?[] sourceColumns = table.Definition.GetFirstChild<TableColumns>()?.Elements<TableColumn>().Select(column => column.Id?.Value).ToArray() ?? Array.Empty<uint?>();
                uint?[] currentColumns = current.GetFirstChild<TableColumns>()?.Elements<TableColumn>().Select(column => column.Id?.Value).ToArray() ?? Array.Empty<uint?>();
                if (!sourceColumns.SequenceEqual(currentColumns)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot add, remove or reorder columns of table '{table.DisplayName}' on worksheet '{context.Sheet.Name}'; formulas refer to table columns by position.");
                }
                replacements[table.PartName] = CreateTablePart(context, current);
                generatedParts.Add(table.PartName);
            }
        }

        /// <summary>Encodes one CT_Table as the TABLE record sequence of a Table part.</summary>
        internal static byte[] CreateTablePart(XlsbWorksheetFeatureContext context, Table table) {
            string sheetName = context.Sheet.Name;
            string label = table.DisplayName?.Value ?? table.Name?.Value ?? "(unnamed)";
            NotSupportedException Refuse(string detail) =>
                new NotSupportedException($"Native XLSB generation cannot encode table '{label}' on worksheet '{sheetName}': {detail}.");

            if (XlsbOpenXmlAttributeValidator.FindUnsupported(table, TableAttributes) is OpenXmlAttribute attribute) {
                throw Refuse($"attribute '{attribute.LocalName}' (differential formats, cell styles and connections are not encoded)");
            }
            if (table.TableType != null && table.TableType.Value != TableValues.Worksheet) {
                throw Refuse("XML-mapped and query tables are not supported");
            }
            if (table.ChildElements.Any(child => child is not AutoFilter && child is not TableColumns && child is not TableStyleInfo)
                || table.Elements<AutoFilter>().Count() > 1 || table.Elements<TableColumns>().Count() != 1 || table.Elements<TableStyleInfo>().Count() > 1) {
                throw Refuse("only autoFilter, tableColumns and tableStyleInfo children are supported (sort state and extensions are not)");
            }

            uint id = table.Id?.Value ?? 0U;
            if (id < 1U || id > 0xFFFFFFFEU) throw Refuse("the table id must be between 1 and 4294967294");
            string? displayName = table.DisplayName?.Value;
            if (string.IsNullOrEmpty(displayName) || displayName!.Length > MaximumNameLength
                || displayName.StartsWith("_xl", StringComparison.OrdinalIgnoreCase)) {
                throw Refuse("the display name must have 1 to 255 characters and not start with _xl");
            }
            string? name = table.Name?.Value;
            if (name != null && name.Length > MaximumNameLength) throw Refuse("the name exceeds 255 characters");
            string? comment = table.Comment?.Value;
            if (comment != null && comment.Length > MaximumNameLength) throw Refuse("the comment exceeds 255 characters");
            if (table.Reference?.Value is not string reference
                || !XlsbPayloadWriter.TryParseRange(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                throw Refuse($"the range '{table.Reference?.Value}' is invalid");
            }
            uint headerRows = table.HeaderRowCount?.Value ?? 1U;
            uint totalsRows = table.TotalsRowCount?.Value ?? 0U;
            if (headerRows > 1U || totalsRows > 1U) throw Refuse("only zero or one header and totals rows are supported");
            if (lastRow - firstRow + 1 <= headerRows + totalsRows) throw Refuse("the range must hold at least one data row");

            TableColumn[] columns = table.TableColumns!.Elements<TableColumn>().ToArray();
            int width = lastColumn - firstColumn + 1;
            if (table.TableColumns.ChildElements.Any(child => child is not TableColumn)
                || XlsbOpenXmlAttributeValidator.FindUnsupported(table.TableColumns, new[] { "count" }) is OpenXmlAttribute
                || columns.Length != width) {
                throw Refuse("the table must define exactly one column per range column");
            }

            var records = new List<XlsbGeneratedRecord>();
            using (var payload = new MemoryStream(96 + (displayName.Length + (name?.Length ?? 0)) * 2)) {
                XlsbPayloadWriter.WriteRfX(payload, firstRow, firstColumn, lastRow, lastColumn);
                XlsbPayloadWriter.WriteUInt32(payload, 0U); // lt: LTRANGE
                XlsbPayloadWriter.WriteUInt32(payload, id);
                XlsbPayloadWriter.WriteUInt32(payload, headerRows);
                XlsbPayloadWriter.WriteUInt32(payload, totalsRows);
                uint flags = 0;
                // CT_Table totalsRowShown defaults to true.
                if (table.TotalsRowShown?.Value ?? true) flags |= 0x01U;
                if (table.InsertRow?.Value == true) flags |= 0x04U;
                if (table.InsertRowShift?.Value == true) flags |= 0x08U;
                if (table.Published?.Value == true) flags |= 0x10U;
                XlsbPayloadWriter.WriteUInt32(payload, flags);
                for (int index = 0; index < 6; index++) XlsbPayloadWriter.WriteUInt32(payload, NoDxf);
                XlsbPayloadWriter.WriteUInt32(payload, 0U); // dwConnID
                XlsbPayloadWriter.WriteNullableWideString(payload, name);
                XlsbPayloadWriter.WriteNullableWideString(payload, displayName);
                XlsbPayloadWriter.WriteNullableWideString(payload, comment);
                for (int index = 0; index < 3; index++) XlsbPayloadWriter.WriteNullableWideString(payload, null);
                records.Add(new XlsbGeneratedRecord(BrtBeginList, payload.ToArray()));
            }

            AutoFilter? autoFilter = table.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                if (headerRows == 0U) throw Refuse("a table without a header row cannot carry an AutoFilter");
                records.AddRange(XlsbWorksheetAutoFilterWriter.CreateRecords(autoFilter, sheetName));
            }

            records.Add(new XlsbGeneratedRecord(BrtBeginListCols, UInt32Payload(checked((uint)columns.Length))));
            var ids = new HashSet<uint>();
            var captions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < columns.Length; index++) {
                records.AddRange(CreateColumn(context, displayName, columns[index], firstColumn + index, firstRow, lastRow, headerRows, totalsRows, ids, captions, Refuse));
            }
            records.Add(new XlsbGeneratedRecord(BrtEndListCols, Array.Empty<byte>()));
            records.Add(new XlsbGeneratedRecord(BrtTableStyleClient, CreateStyle(table.GetFirstChild<TableStyleInfo>(), Refuse)));
            records.Add(new XlsbGeneratedRecord(BrtEndList, Array.Empty<byte>()));

            using var output = new MemoryStream();
            foreach (XlsbGeneratedRecord record in records) Biff12.XlsbRecordWriter.Write(output, record.Type, record.Payload);
            return output.ToArray();
        }

        private static IEnumerable<XlsbGeneratedRecord> CreateColumn(XlsbWorksheetFeatureContext context, string tableName, TableColumn column, int sheetColumn,
            int firstRow, int lastRow, uint headerRows, uint totalsRows, HashSet<uint> ids, HashSet<string> captions,
            Func<string, NotSupportedException> refuse) {
            if (XlsbOpenXmlAttributeValidator.FindUnsupported(column, ColumnAttributes) is OpenXmlAttribute attribute) {
                throw refuse($"column attribute '{attribute.LocalName}' is not supported");
            }
            if (column.ChildElements.Any(child => child is not CalculatedColumnFormula && child is not TotalsRowFormula)) {
                throw refuse("table column XML mappings and extensions are not supported");
            }
            uint id = column.Id?.Value ?? 0U;
            if (id < 1U || !ids.Add(id)) throw refuse("column ids must be unique and at least 1");
            string caption = column.Name?.Value ?? string.Empty;
            if (caption.Length < 1 || caption.Length > MaximumNameLength || !captions.Add(caption)) {
                throw refuse("column names must be unique and have 1 to 255 characters");
            }
            if (headerRows == 1U && !string.Equals(CellText(context.Sheet, firstRow, sheetColumn), caption, StringComparison.Ordinal)) {
                throw refuse($"the header cell of column '{caption}' does not hold the column name");
            }

            TotalsRowFunctionValues function = column.TotalsRowFunction?.Value ?? TotalsRowFunctionValues.None;
            uint ilta = TotalsFunctions.First(entry => entry.Function == function).Value;
            string? total = column.TotalsRowLabel?.Value;
            TotalsRowFormula? totalsFormula = column.GetFirstChild<TotalsRowFormula>();
            if (ilta == 9U ? total != null || totalsFormula == null : totalsFormula != null) {
                throw refuse($"column '{caption}' must pair a custom totals function with exactly one totals formula and no label");
            }
            if (total != null) {
                if (total.Length > MaximumTotalLabelLength) throw refuse("a totals label exceeds 8189 characters");
                if (totalsRows == 1U && !string.Equals(CellText(context.Sheet, lastRow, sheetColumn), total, StringComparison.Ordinal)) {
                    throw refuse($"the totals cell of column '{caption}' does not hold its label");
                }
            }

            var records = new List<XlsbGeneratedRecord>(4);
            using (var payload = new MemoryStream(40 + caption.Length * 2)) {
                XlsbPayloadWriter.WriteUInt32(payload, id);
                XlsbPayloadWriter.WriteUInt32(payload, ilta);
                XlsbPayloadWriter.WriteUInt32(payload, NoDxf); // nDxfHdr
                XlsbPayloadWriter.WriteUInt32(payload, NoDxf); // nDxfInsertRow
                XlsbPayloadWriter.WriteUInt32(payload, NoDxf); // nDxfAgg
                XlsbPayloadWriter.WriteUInt32(payload, 0U);    // idqsif
                XlsbPayloadWriter.WriteNullableWideString(payload, null); // stName is NULL for LTRANGE tables
                XlsbPayloadWriter.WriteNullableWideString(payload, caption);
                XlsbPayloadWriter.WriteNullableWideString(payload, total);
                for (int index = 0; index < 3; index++) XlsbPayloadWriter.WriteNullableWideString(payload, null);
                records.Add(new XlsbGeneratedRecord(BrtBeginListCol, payload.ToArray()));
            }
            CalculatedColumnFormula? calculated = column.GetFirstChild<CalculatedColumnFormula>();
            if (calculated != null) records.Add(CreateFormulaRecord(context, tableName, BrtListCCFmla, calculated, $"calculated column '{caption}'", refuse));
            if (totalsFormula != null) records.Add(CreateFormulaRecord(context, tableName, BrtListTrFmla, totalsFormula, $"totals row of column '{caption}'", refuse));
            records.Add(new XlsbGeneratedRecord(BrtEndListCol, Array.Empty<byte>()));
            return records;
        }

        private static XlsbGeneratedRecord CreateFormulaRecord(XlsbWorksheetFeatureContext context, string tableName, int type, TableFormulaType formula, string purpose,
            Func<string, NotSupportedException> refuse) {
            if (formula.HasChildren || XlsbOpenXmlAttributeValidator.FindUnsupported(formula, new[] { "array" }) is OpenXmlAttribute) {
                throw refuse($"the {purpose} formula carries unsupported attributes");
            }
            bool array = formula.Array?.Value == true;
            // Inside a table definition an unqualified reference such as [@Amount] belongs to this table; PtgList names it.
            string text = formula.Text ?? string.Empty;
            if (text.IndexOf('[') >= 0) {
                try {
                    text = ExcelFormulaSyntaxTree.Parse(text).RewriteStructuredReferences((name, selector) => (name ?? tableName) + selector);
                } catch (ArgumentException) {
                    // The encoder reports formulas it cannot parse.
                }
            }
            byte[] encoded = context.EncodeFormula(text, "table " + purpose,
                array ? ExcelNativeFormulaOperandTypes.Array : ExcelNativeFormulaOperandTypes.Value);
            var payload = new byte[1 + encoded.Length];
            payload[0] = array ? (byte)0x02 : (byte)0x00; // fArray is bit 1
            Buffer.BlockCopy(encoded, 0, payload, 1, encoded.Length);
            return new XlsbGeneratedRecord(type, payload);
        }

        private static byte[] CreateStyle(TableStyleInfo? style, Func<string, NotSupportedException> refuse) {
            string? name = null;
            ushort flags = 0;
            if (style != null) {
                if (style.HasChildren || XlsbOpenXmlAttributeValidator.FindUnsupported(style, StyleAttributes) is OpenXmlAttribute) {
                    throw refuse("the table style info carries unsupported content");
                }
                name = style.Name?.Value;
                if (name != null && !IsBuiltInTableStyle(name)) {
                    throw refuse($"table style '{name}' is not a built-in table style (custom table styles are not encoded)");
                }
                if (style.ShowFirstColumn?.Value == true) flags |= 0x01;
                if (style.ShowLastColumn?.Value == true) flags |= 0x02;
                if (style.ShowRowStripes?.Value == true) flags |= 0x04;
                if (style.ShowColumnStripes?.Value == true) flags |= 0x08;
            }
            using var payload = new MemoryStream(6 + (name?.Length ?? 0) * 2);
            XlsbPayloadWriter.WriteUInt16(payload, flags);
            XlsbPayloadWriter.WriteNullableWideString(payload, name);
            return payload.ToArray();
        }

        internal static bool IsBuiltInTableStyle(string name) =>
            name.StartsWith("TableStyle", StringComparison.Ordinal)
            && Enum.TryParse(name, ignoreCase: false, out ExcelTableStyle style)
            && style != ExcelTableStyle.None
            && string.Equals(style.ToString(), name, StringComparison.Ordinal);

        private static string CellText(ExcelSheet sheet, int row, int column) {
            Cell? cell = sheet.TryGetExistingCell(row, column);
            return cell == null ? string.Empty : sheet.GetCellText(cell);
        }

        private static bool IsRelationshipId(OpenXmlAttribute attribute) =>
            attribute.LocalName == "id"
            && attribute.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static byte[] UInt32Payload(uint value) {
            using var payload = new MemoryStream(4);
            XlsbPayloadWriter.WriteUInt32(payload, value);
            return payload.ToArray();
        }
    }
}
