using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Writes a DataSet directly to an XLSX package, using one worksheet and one Excel table per DataTable.
        /// This path is intended for export workloads where the caller does not need to keep editing the workbook object.
        /// </summary>
        public static IReadOnlyList<ExcelDataSetImportResult> WriteDataSet(
            Stream stream,
            DataSet dataSet,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeHeaders = true,
            bool includeAutoFilter = true,
            CancellationToken ct = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (dataSet == null) throw new ArgumentNullException(nameof(dataSet));
            if (dataSet.Tables.Count == 0) throw new ArgumentException("The DataSet must contain at least one DataTable.", nameof(dataSet));

            var model = DirectDataSetWorkbookModel.Create(dataSet, tableStyle, includeHeaders, includeAutoFilter, ct);
            if (stream.CanSeek) {
                PrepareDestinationStreamForWrite(stream);
            }

            DirectDataSetWorkbookWriter.Write(stream, model, ct);
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            return model.Results;
        }

        private sealed class DirectDataSetWorkbookModel {
            private DirectDataSetWorkbookModel(IReadOnlyList<DirectDataSetSheetModel> sheets, IReadOnlyList<ExcelDataSetImportResult> results) {
                Sheets = sheets;
                Results = results;
            }

            internal IReadOnlyList<DirectDataSetSheetModel> Sheets { get; }

            internal IReadOnlyList<ExcelDataSetImportResult> Results { get; }

            internal static DirectDataSetWorkbookModel Create(
                DataSet dataSet,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                CancellationToken ct) {
                var sheets = new List<DirectDataSetSheetModel>();
                var results = new List<ExcelDataSetImportResult>();
                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int index = 1;
                foreach (DataTable table in dataSet.Tables) {
                    ct.ThrowIfCancellationRequested();
                    string requestedName = string.IsNullOrWhiteSpace(table.TableName)
                        ? "Table" + index.ToString(CultureInfo.InvariantCulture)
                        : table.TableName;
                    string sheetName = GetUniqueName(SanitizeSheetName(requestedName), usedSheetNames, 31);
                    string tableName = GetUniqueName(SanitizeTableName(requestedName), usedTableNames, 255);
                    int rowCount = table.Rows.Count + (includeHeaders ? 1 : 0);
                    string range = table.Columns.Count == 0 || rowCount == 0
                        ? string.Empty
                        : "A1:" + A1.CellReference(rowCount, table.Columns.Count);

                    var sheet = new DirectDataSetSheetModel(index, sheetName, tableName, range, table, tableStyle, includeHeaders, includeAutoFilter);
                    sheets.Add(sheet);
                    results.Add(new ExcelDataSetImportResult(sheetName, range.Length == 0 ? null : tableName, range, table.Rows.Count, table.Columns.Count));
                    index++;
                }

                return new DirectDataSetWorkbookModel(sheets, results);
            }

            private static string GetUniqueName(string baseName, HashSet<string> used, int maxLength) {
                string trimmed = string.IsNullOrWhiteSpace(baseName) ? "Table" : baseName;
                if (trimmed.Length > maxLength) {
                    trimmed = trimmed.Substring(0, maxLength);
                }

                if (used.Add(trimmed)) {
                    return trimmed;
                }

                int suffix = 2;
                while (true) {
                    string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                    int prefixLength = Math.Max(1, maxLength - suffixText.Length);
                    string candidate = trimmed.Length > prefixLength
                        ? trimmed.Substring(0, prefixLength) + suffixText
                        : trimmed + suffixText;
                    if (used.Add(candidate)) {
                        return candidate;
                    }

                    suffix++;
                }
            }

            private static string SanitizeSheetName(string name) {
                var builder = new StringBuilder(name.Length);
                foreach (char ch in name) {
                    builder.Append(ch is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ? '_' : ch);
                }

                string value = builder.ToString().Trim('\'');
                return string.IsNullOrWhiteSpace(value) ? "Table" : value;
            }

            private static string SanitizeTableName(string name) {
                var builder = new StringBuilder(name.Length + 1);
                foreach (char ch in name) {
                    builder.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
                }

                string value = builder.ToString();
                if (string.IsNullOrWhiteSpace(value)) {
                    value = "Table";
                }

                if (!char.IsLetter(value[0]) && value[0] != '_') {
                    value = "_" + value;
                }

                return value;
            }
        }

        private sealed class DirectDataSetSheetModel {
            internal DirectDataSetSheetModel(
                int index,
                string sheetName,
                string tableName,
                string range,
                DataTable table,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter) {
                Index = index;
                SheetName = sheetName;
                TableName = tableName;
                Range = range;
                Table = table;
                TableStyle = tableStyle;
                IncludeHeaders = includeHeaders;
                IncludeAutoFilter = includeAutoFilter;
            }

            internal int Index { get; }

            internal string SheetName { get; }

            internal string TableName { get; }

            internal string Range { get; }

            internal DataTable Table { get; }

            internal TableStyle TableStyle { get; }

            internal bool IncludeHeaders { get; }

            internal bool IncludeAutoFilter { get; }
        }

        private static class DirectDataSetWorkbookWriter {
            private const long MaxSafeInteger = 9007199254740991L;
            private const ulong MaxSafeUnsignedInteger = 9007199254740991UL;

            internal static void Write(Stream stream, DirectDataSetWorkbookModel model, CancellationToken ct) {
                using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
                WriteContentTypes(archive, model.Sheets);
                WriteTextEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                WriteWorkbook(archive, model.Sheets);
                WriteWorkbookRelationships(archive, model.Sheets.Count);
                WriteStyles(archive);
                foreach (var sheet in model.Sheets) {
                    ct.ThrowIfCancellationRequested();
                    WriteWorksheet(archive, sheet, ct);
                    if (sheet.Range.Length > 0) {
                        WriteTextEntry(archive, $"xl/worksheets/_rels/sheet{sheet.Index}.xml.rels",
                            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                            $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{sheet.Index}.xml\"/>" +
                            "</Relationships>");
                        WriteTable(archive, sheet);
                    }
                }
            }

            private static void WriteContentTypes(ZipArchive archive, IReadOnlyList<DirectDataSetSheetModel> sheets) {
                var builder = new StringBuilder(1024 + sheets.Count * 260);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
                builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                foreach (var sheet in sheets) {
                    builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                    builder.Append(sheet.Index.ToString(CultureInfo.InvariantCulture));
                    builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                    if (sheet.Range.Length > 0) {
                        builder.Append("<Override PartName=\"/xl/tables/table");
                        builder.Append(sheet.Index.ToString(CultureInfo.InvariantCulture));
                        builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
                    }
                }

                builder.Append("</Types>");
                WriteTextEntry(archive, "[Content_Types].xml", builder.ToString());
            }

            private static void WriteWorkbook(ZipArchive archive, IReadOnlyList<DirectDataSetSheetModel> sheets) {
                var builder = new StringBuilder(256 + sheets.Count * 120);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
                foreach (var sheet in sheets) {
                    builder.Append("<sheet name=\"");
                    AppendEscaped(builder, sheet.SheetName);
                    builder.Append("\" sheetId=\"");
                    builder.Append(sheet.Index.ToString(CultureInfo.InvariantCulture));
                    builder.Append("\" r:id=\"rId");
                    builder.Append(sheet.Index.ToString(CultureInfo.InvariantCulture));
                    builder.Append("\"/>");
                }

                builder.Append("</sheets></workbook>");
                WriteTextEntry(archive, "xl/workbook.xml", builder.ToString());
            }

            private static void WriteWorkbookRelationships(ZipArchive archive, int sheetCount) {
                var builder = new StringBuilder(384 + sheetCount * 160);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 1; i <= sheetCount; i++) {
                    builder.Append("<Relationship Id=\"rId");
                    builder.Append(i.ToString(CultureInfo.InvariantCulture));
                    builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
                    builder.Append(i.ToString(CultureInfo.InvariantCulture));
                    builder.Append(".xml\"/>");
                }

                builder.Append("<Relationship Id=\"rId");
                builder.Append((sheetCount + 1).ToString(CultureInfo.InvariantCulture));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                builder.Append("</Relationships>");
                WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", builder.ToString());
            }

            private static void WriteStyles(ZipArchive archive) {
                WriteTextEntry(archive, "xl/styles.xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    "<numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd hh:mm\"/><numFmt numFmtId=\"165\" formatCode=\"[h]:mm:ss\"/></numFmts>" +
                    "<fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts>" +
                    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                    "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>" +
                    "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                    "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>" +
                    "</styleSheet>");
            }

            private static void WriteWorksheet(ZipArchive archive, DirectDataSetSheetModel sheet, CancellationToken ct) {
                var builder = new StringBuilder(Math.Max(4096, (sheet.Table.Rows.Count + 1) * Math.Max(1, sheet.Table.Columns.Count) * 40));
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                builder.Append("<dimension ref=\"");
                AppendEscaped(builder, sheet.Range.Length == 0 ? "A1" : sheet.Range);
                builder.Append("\"/><sheetData>");
                int rowIndex = 1;
                if (sheet.IncludeHeaders) {
                    builder.Append("<row r=\"1\">");
                    for (int c = 0; c < sheet.Table.Columns.Count; c++) {
                        AppendCell(builder, 1, c + 1, sheet.Table.Columns[c].ColumnName, null);
                    }

                    builder.Append("</row>");
                    rowIndex++;
                }

                foreach (DataRow row in sheet.Table.Rows) {
                    ct.ThrowIfCancellationRequested();
                    builder.Append("<row r=\"");
                    builder.Append(rowIndex.ToString(CultureInfo.InvariantCulture));
                    builder.Append("\">");
                    for (int c = 0; c < sheet.Table.Columns.Count; c++) {
                        object? value = row.IsNull(c) ? null : row[c];
                        AppendCell(builder, rowIndex, c + 1, value, GetStyleIndex(sheet.Table.Columns[c].DataType));
                    }

                    builder.Append("</row>");
                    rowIndex++;
                }

                builder.Append("</sheetData>");
                if (sheet.Range.Length > 0) {
                    builder.Append("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>");
                }

                builder.Append("</worksheet>");
                WriteTextEntry(archive, $"xl/worksheets/sheet{sheet.Index}.xml", builder.ToString());
            }

            private static void WriteTable(ZipArchive archive, DirectDataSetSheetModel sheet) {
                var builder = new StringBuilder(512 + sheet.Table.Columns.Count * 80);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
                builder.Append(sheet.Index.ToString(CultureInfo.InvariantCulture));
                builder.Append("\" name=\"");
                AppendEscaped(builder, sheet.TableName);
                builder.Append("\" displayName=\"");
                AppendEscaped(builder, sheet.TableName);
                builder.Append("\" ref=\"");
                AppendEscaped(builder, sheet.Range);
                builder.Append("\" headerRowCount=\"");
                builder.Append(sheet.IncludeHeaders ? "1" : "0");
                builder.Append("\" totalsRowShown=\"0\">");
                if (sheet.IncludeAutoFilter && sheet.IncludeHeaders) {
                    builder.Append("<autoFilter ref=\"");
                    AppendEscaped(builder, sheet.Range);
                    builder.Append("\"/>");
                }

                builder.Append("<tableColumns count=\"");
                builder.Append(sheet.Table.Columns.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append("\">");
                for (int i = 0; i < sheet.Table.Columns.Count; i++) {
                    builder.Append("<tableColumn id=\"");
                    builder.Append((i + 1).ToString(CultureInfo.InvariantCulture));
                    builder.Append("\" name=\"");
                    AppendEscaped(builder, string.IsNullOrWhiteSpace(sheet.Table.Columns[i].ColumnName) ? "Column" + (i + 1).ToString(CultureInfo.InvariantCulture) : sheet.Table.Columns[i].ColumnName);
                    builder.Append("\"/>");
                }

                builder.Append("</tableColumns><tableStyleInfo name=\"");
                builder.Append(sheet.TableStyle.ToString());
                builder.Append("\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>");
                WriteTextEntry(archive, $"xl/tables/table{sheet.Index}.xml", builder.ToString());
            }

            private static uint? GetStyleIndex(Type dataType) {
                if (dataType == typeof(DateTime) || dataType == typeof(DateTimeOffset)) return 1U;
                if (dataType == typeof(TimeSpan)) return 2U;
                return null;
            }

            private static void AppendCell(StringBuilder builder, int row, int column, object? value, uint? styleIndex) {
                builder.Append("<c r=\"");
                builder.Append(A1.CellReference(row, column));
                builder.Append('"');
                if (styleIndex.HasValue) {
                    builder.Append(" s=\"");
                    builder.Append(styleIndex.Value.ToString(CultureInfo.InvariantCulture));
                    builder.Append('"');
                }

                string text;
                switch (value) {
                    case null:
                    case DBNull:
                        builder.Append(" t=\"str\"><v></v></c>");
                        return;
                    case string stringValue:
                        AppendStringCell(builder, stringValue);
                        return;
                    case bool boolValue:
                        builder.Append(" t=\"b\"><v>");
                        builder.Append(boolValue ? "1" : "0");
                        builder.Append("</v></c>");
                        return;
                    case DateTime dateTime:
                        text = dateTime.ToOADate().ToString(CultureInfo.InvariantCulture);
                        break;
                    case DateTimeOffset dateTimeOffset:
                        text = dateTimeOffset.LocalDateTime.ToOADate().ToString(CultureInfo.InvariantCulture);
                        break;
                    case TimeSpan timeSpan:
                        text = timeSpan.TotalDays.ToString(CultureInfo.InvariantCulture);
                        break;
                    case double doubleValue:
                        text = doubleValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case float floatValue:
                        text = floatValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case decimal decimalValue:
                        text = decimalValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case sbyte sbyteValue:
                        text = sbyteValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case byte byteValue:
                        text = byteValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case short shortValue:
                        text = shortValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case ushort ushortValue:
                        text = ushortValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case int intValue:
                        text = intValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case uint uintValue:
                        text = uintValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case long longValue when longValue >= -MaxSafeInteger && longValue <= MaxSafeInteger:
                        text = longValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case ulong ulongValue when ulongValue <= MaxSafeUnsignedInteger:
                        text = ulongValue.ToString(CultureInfo.InvariantCulture);
                        break;
#if NET6_0_OR_GREATER
                    case DateOnly dateOnly:
                        text = dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture);
                        break;
                    case TimeOnly timeOnly:
                        text = timeOnly.ToTimeSpan().TotalDays.ToString(CultureInfo.InvariantCulture);
                        break;
#endif
                    default:
                        AppendStringCell(builder, value.ToString() ?? string.Empty);
                        return;
                }

                builder.Append("><v>");
                AppendEscaped(builder, text);
                builder.Append("</v></c>");
            }

            private static void AppendStringCell(StringBuilder builder, string text) {
                CoerceValueHelper.ValidateSharedStringLength(text, "value");
                builder.Append(" t=\"str\"><v>");
                AppendEscaped(builder, Utilities.ExcelSanitizer.SanitizeString(text));
                builder.Append("</v></c>");
            }

            private static void WriteTextEntry(ZipArchive archive, string path, string text) {
                var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                writer.Write(text);
            }

            private static void AppendEscaped(StringBuilder builder, string value) {
                foreach (char ch in value) {
                    switch (ch) {
                        case '&':
                            builder.Append("&amp;");
                            break;
                        case '<':
                            builder.Append("&lt;");
                            break;
                        case '>':
                            builder.Append("&gt;");
                            break;
                        case '"':
                            builder.Append("&quot;");
                            break;
                        case '\'':
                            builder.Append("&apos;");
                            break;
                        default:
                            builder.Append(ch);
                            break;
                    }
                }
            }
        }
    }
}
