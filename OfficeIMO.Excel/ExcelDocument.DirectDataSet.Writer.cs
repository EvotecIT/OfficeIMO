using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static class DirectDataSetWorkbookWriter {
            private const long MaxSafeInteger = 9007199254740991L;
            private const ulong MaxSafeUnsignedInteger = 9007199254740991UL;
            private const int XmlWriterBufferSize = 65536;
            private const string DateStyleAttribute = " s=\"1\"";
            private const string TimeStyleAttribute = " s=\"2\"";
            private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            private static readonly DateTime ExcelDateTimeOffsetEpoch = DateTime.FromOADate(0);

            internal static void Write(Stream stream, DirectDataSetWorkbookModel model, CancellationToken ct) {
                using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
                WriteContentTypes(archive, model.Sheets);
                WriteTextEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                    "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                    "</Relationships>");
                WriteCoreProperties(archive);
                WriteAppProperties(archive);
                WriteWorkbook(archive, model.Sheets);
                WriteWorkbookRelationships(archive, model.Sheets.Count);
                WriteStyles(archive);
                foreach (var sheet in model.Sheets) {
                    ct.ThrowIfCancellationRequested();
                    WriteWorksheet(archive, sheet, model.DateTimeOffsetWriteStrategy, ct);
                    if (sheet.HasTable) {
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
                builder.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
                builder.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
                builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                foreach (var sheet in sheets) {
                    builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                    builder.Append(sheet.Index.ToString(CultureInfo.InvariantCulture));
                    builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                    if (sheet.HasTable) {
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

            private static void WriteCoreProperties(ZipArchive archive) {
                WriteTextEntry(archive, "docProps/core.xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                    "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
                    "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                    "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
                    "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>");
            }

            private static void WriteAppProperties(ZipArchive archive) {
                WriteTextEntry(archive, "docProps/app.xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                    "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                    "<Application>OfficeIMO.Excel</Application>" +
                    "</Properties>");
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

            private static void WriteWorksheet(ZipArchive archive, DirectDataSetSheetModel sheet, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, CancellationToken ct) {
                var entry = archive.CreateEntry($"xl/worksheets/sheet{sheet.Index}.xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                writer.Write("<dimension ref=\"");
                WriteEscaped(writer, sheet.Range.Length == 0 ? "A1" : sheet.Range);
                writer.Write("\"/>");
                WriteColumns(writer, sheet.ColumnWidths);
                writer.Write("<sheetData>");
                int columnCount = sheet.Table.ColumnCount;
                int rowCount = sheet.Table.RowCount;
                string[] cellReferencePrefixes = CreateCellReferencePrefixes(columnCount);
                string?[]? styleAttributes = CreateStyleAttributes(sheet.Table);
                bool[]? valueStyleColumns = CreateValueStyleColumns(sheet.Table);
                int rowIndex = 1;
                if (sheet.IncludeHeaders) {
                    const string headerRowReference = "1";
                    writer.Write("<row r=\"1\">");
                    for (int c = 0; c < columnCount; c++) {
                        WriteCell(writer, headerRowReference, cellReferencePrefixes[c], sheet.Table.GetColumnName(c), null, dateTimeOffsetWriteStrategy);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }

                if (valueStyleColumns == null) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        ct.ThrowIfCancellationRequested();
                        bool rowStarted = false;
                        string rowReference = rowIndex.ToString(CultureInfo.InvariantCulture);
                        var sourceRow = sheet.Table.GetRow(sourceRowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = sourceRow.GetValue(c);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], dateTimeOffsetWriteStrategy);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }
                } else {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        ct.ThrowIfCancellationRequested();
                        bool rowStarted = false;
                        string rowReference = rowIndex.ToString(CultureInfo.InvariantCulture);
                        var sourceRow = sheet.Table.GetRow(sourceRowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = sourceRow.GetValue(c);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns[c], dateTimeOffsetWriteStrategy);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }
                }

                writer.Write("</sheetData>");
                if (sheet.HasTable) {
                    writer.Write("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>");
                }

                writer.Write("</worksheet>");
            }

            private static string[] CreateCellReferencePrefixes(int columnCount) {
                var columns = new string[columnCount];
                for (int i = 0; i < columnCount; i++) {
                    columns[i] = "<c r=\"" + A1.ColumnIndexToLetters(i + 1);
                }

                return columns;
            }

            private static string?[]? CreateStyleAttributes(DirectDataSetTableModel table) {
                string?[]? styleAttributes = null;
                for (int i = 0; i < table.ColumnCount; i++) {
                    string? styleAttribute = CreateStyleAttribute(table.GetColumnType(i));
                    if (styleAttribute == null) {
                        continue;
                    }

                    styleAttributes ??= new string?[table.ColumnCount];
                    styleAttributes[i] = styleAttribute;
                }

                return styleAttributes;
            }

            private static bool[]? CreateValueStyleColumns(DirectDataSetTableModel table) {
                bool[]? valueStyleColumns = null;
                for (int i = 0; i < table.ColumnCount; i++) {
                    if (table.GetColumnType(i) != typeof(object)) {
                        continue;
                    }

                    valueStyleColumns ??= new bool[table.ColumnCount];
                    valueStyleColumns[i] = true;
                }

                return valueStyleColumns;
            }

            private static void WriteColumns(TextWriter writer, double[]? columnWidths) {
                if (columnWidths == null || columnWidths.Length == 0) {
                    return;
                }

                bool started = false;
                for (int i = 0; i < columnWidths.Length; i++) {
                    double width = NormalizeColumnWidth(columnWidths[i]);
                    if (width <= 0D) {
                        continue;
                    }

                    if (!started) {
                        writer.Write("<cols>");
                        started = true;
                    }

                    writer.Write("<col min=\"");
                    WriteInvariant(writer, i + 1);
                    writer.Write("\" max=\"");
                    WriteInvariant(writer, i + 1);
                    writer.Write("\" width=\"");
                    WriteColumnWidth(writer, width);
                    writer.Write("\" bestFit=\"1\" customWidth=\"1\"/>");
                }

                if (started) {
                    writer.Write("</cols>");
                }
            }

            private static double NormalizeColumnWidth(double width) {
                if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) {
                    return 0D;
                }

                return Math.Min(width, 255D);
            }

            private static void WriteTable(ZipArchive archive, DirectDataSetSheetModel sheet) {
                var entry = archive.CreateEntry($"xl/tables/table{sheet.Index}.xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
                WriteInvariant(writer, sheet.Index);
                writer.Write("\" name=\"");
                WriteEscaped(writer, sheet.TableName!);
                writer.Write("\" displayName=\"");
                WriteEscaped(writer, sheet.TableName!);
                writer.Write("\" ref=\"");
                WriteEscaped(writer, sheet.Range);
                writer.Write("\" headerRowCount=\"");
                writer.Write(sheet.IncludeHeaders ? "1" : "0");
                writer.Write("\" totalsRowShown=\"0\">");
                if (sheet.IncludeAutoFilter && sheet.IncludeHeaders) {
                    writer.Write("<autoFilter ref=\"");
                    WriteEscaped(writer, sheet.Range);
                    writer.Write("\"/>");
                }

                writer.Write("<tableColumns count=\"");
                int columnCount = sheet.Table.ColumnCount;
                WriteInvariant(writer, columnCount);
                writer.Write("\">");
                for (int i = 0; i < columnCount; i++) {
                    writer.Write("<tableColumn id=\"");
                    WriteInvariant(writer, i + 1);
                    writer.Write("\" name=\"");
                    string columnName = sheet.Table.GetColumnName(i);
                    if (string.IsNullOrWhiteSpace(columnName)) {
                        writer.Write("Column");
                        WriteInvariant(writer, i + 1);
                    } else {
                        WriteEscaped(writer, columnName);
                    }

                    writer.Write("\"/>");
                }

                writer.Write("</tableColumns><tableStyleInfo name=\"");
                writer.Write(sheet.TableStyle.ToString());
                writer.Write("\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>");
            }

            private static uint? GetStyleIndex(Type dataType) {
                if (dataType == typeof(DateTime) || dataType == typeof(DateTimeOffset)) return 1U;
                if (dataType == typeof(TimeSpan)) return 2U;
#if NET6_0_OR_GREATER
                if (dataType == typeof(DateOnly)) return 1U;
                if (dataType == typeof(TimeOnly)) return 2U;
#endif
                return null;
            }

            private static string? CreateStyleAttribute(Type dataType) {
                uint? styleIndex = GetStyleIndex(dataType);
                return styleIndex switch {
                    1U => DateStyleAttribute,
                    2U => TimeStyleAttribute,
                    _ => null
                };
            }

            private static string? CreateStyleAttributeForValue(object? value) {
                if (value == null || value == DBNull.Value) {
                    return null;
                }

                return CreateStyleAttribute(value.GetType());
            }

            private static bool IsBlankCellValue(object? value) => value == null || value == DBNull.Value;

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                if (styleAttribute != null) {
                    writer.Write(styleAttribute);
                }

                WriteCellValue(writer, value, dateTimeOffsetWriteStrategy);
            }

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, bool useValueStyle, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                string? effectiveStyleAttribute = styleAttribute ?? (useValueStyle ? CreateStyleAttributeForValue(value) : null);
                if (effectiveStyleAttribute != null) {
                    writer.Write(effectiveStyleAttribute);
                }

                WriteCellValue(writer, value, dateTimeOffsetWriteStrategy);
            }

            private static void WriteCellValue(TextWriter writer, object? value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                switch (value) {
                    case null:
                    case DBNull:
                        writer.Write(" t=\"str\"><v></v></c>");
                        return;
                    case string stringValue:
                        WriteStringCell(writer, stringValue);
                        return;
                    case bool boolValue:
                        writer.Write(" t=\"b\"><v>");
                        writer.Write(boolValue ? "1" : "0");
                        writer.Write("</v></c>");
                        return;
                    case DateTime dateTime:
                        WriteRawValueCell(writer, dateTime.ToOADate());
                        return;
                    case DateTimeOffset dateTimeOffset:
                        if (!TryGetDateTimeOffsetSerial(dateTimeOffset, dateTimeOffsetWriteStrategy, out double dateTimeOffsetSerial)) {
                            WriteStringCell(writer, dateTimeOffset.ToString("o", CultureInfo.InvariantCulture));
                            return;
                        }

                        WriteRawValueCell(writer, dateTimeOffsetSerial);
                        return;
                    case TimeSpan timeSpan:
                        WriteRawValueCell(writer, timeSpan.TotalDays);
                        return;
                    case double doubleValue:
                        WriteRawValueCell(writer, doubleValue);
                        return;
                    case float floatValue:
                        WriteRawValueCell(writer, floatValue);
                        return;
                    case decimal decimalValue:
                        WriteRawValueCell(writer, decimalValue);
                        return;
                    case sbyte sbyteValue:
                        WriteRawValueCell(writer, sbyteValue);
                        return;
                    case byte byteValue:
                        WriteRawValueCell(writer, byteValue);
                        return;
                    case short shortValue:
                        WriteRawValueCell(writer, shortValue);
                        return;
                    case ushort ushortValue:
                        WriteRawValueCell(writer, ushortValue);
                        return;
                    case int intValue:
                        WriteRawValueCell(writer, intValue);
                        return;
                    case uint uintValue:
                        WriteRawValueCell(writer, uintValue);
                        return;
                    case long longValue when longValue >= -MaxSafeInteger && longValue <= MaxSafeInteger:
                        WriteRawValueCell(writer, longValue);
                        return;
                    case ulong ulongValue when ulongValue <= MaxSafeUnsignedInteger:
                        WriteRawValueCell(writer, ulongValue);
                        return;
#if NET6_0_OR_GREATER
                    case DateOnly dateOnly:
                        WriteRawValueCell(writer, dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate());
                        return;
                    case TimeOnly timeOnly:
                        WriteRawValueCell(writer, timeOnly.ToTimeSpan().TotalDays);
                        return;
#endif
                    default:
                        WriteStringCell(writer, value.ToString() ?? string.Empty);
                        return;
                }
            }

            private static bool TryGetDateTimeOffsetSerial(DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, out double serial) {
                try {
                    if (value.UtcDateTime < ExcelDateTimeOffsetEpoch) {
                        serial = 0D;
                        return false;
                    }

                    serial = dateTimeOffsetWriteStrategy(value).ToOADate();
                    return true;
                } catch (ArgumentException) {
                    serial = 0D;
                    return false;
                } catch (OverflowException) {
                    serial = 0D;
                    return false;
                }
            }

            private static void WriteStringCell(TextWriter writer, string text) {
                CoerceValueHelper.ValidateSharedStringLength(text, "value");
                writer.Write(" t=\"str\"><v>");
                WriteSanitizedEscaped(writer, text);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, double value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, float value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, decimal value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, long value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, ulong value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteInvariant(TextWriter writer, double value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteColumnWidth(TextWriter writer, double value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, "0.###", CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString("0.###", CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, float value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, decimal value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, int value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[16];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, long value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, ulong value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteTextEntry(ZipArchive archive, string path, string text) {
                var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);
                writer.Write(text);
            }

            private static void AppendEscaped(StringBuilder builder, string value) {
                int escapeIndex = IndexOfXmlEscape(value);
                if (escapeIndex < 0) {
                    builder.Append(value);
                    return;
                }

                int start = 0;
                while (escapeIndex >= 0) {
                    if (escapeIndex > start) {
                        builder.Append(value, start, escapeIndex - start);
                    }

                    AppendEscapedCharacter(builder, value[escapeIndex]);
                    start = escapeIndex + 1;
                    escapeIndex = IndexOfXmlEscape(value, start);
                }

                if (start < value.Length) {
                    builder.Append(value, start, value.Length - start);
                }
            }

            private static void WriteEscaped(TextWriter writer, string value) {
                int escapeIndex = IndexOfXmlEscape(value);
                if (escapeIndex < 0) {
                    writer.Write(value);
                    return;
                }

                int start = 0;
                while (escapeIndex >= 0) {
                    if (escapeIndex > start) {
                        WriteSlice(writer, value, start, escapeIndex - start);
                    }

                    WriteEscapedCharacter(writer, value[escapeIndex]);
                    start = escapeIndex + 1;
                    escapeIndex = IndexOfXmlEscape(value, start);
                }

                if (start < value.Length) {
                    WriteSlice(writer, value, start, value.Length - start);
                }
            }

            private static void WriteSanitizedEscaped(TextWriter writer, string value) {
                int start = 0;
                for (int i = 0; i < value.Length; i++) {
                    char current = value[i];
                    if (!IsInvalidXmlControl(current) && !IsXmlEscape(current)) {
                        continue;
                    }

                    if (i > start) {
                        WriteSlice(writer, value, start, i - start);
                    }

                    if (!IsInvalidXmlControl(current)) {
                        WriteEscapedCharacter(writer, current);
                    }

                    start = i + 1;
                }

                if (start < value.Length) {
                    WriteSlice(writer, value, start, value.Length - start);
                }
            }

            private static void WriteSlice(TextWriter writer, string value, int startIndex, int length) {
#if NET6_0_OR_GREATER
                writer.Write(value.AsSpan(startIndex, length));
#else
                writer.Write(value.Substring(startIndex, length));
#endif
            }

            private static int IndexOfXmlEscape(string value, int startIndex = 0) {
                for (int i = startIndex; i < value.Length; i++) {
                    if (IsXmlEscape(value[i])) {
                        return i;
                    }
                }

                return -1;
            }

            private static bool IsInvalidXmlControl(char value)
                => value < 0x20 && value != '\t' && value != '\n' && value != '\r';

            private static bool IsXmlEscape(char value)
                => value is '&' or '<' or '>' or '"' or '\'';

            private static void AppendEscapedCharacter(StringBuilder builder, char value) {
                switch (value) {
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
                }
            }

            private static void WriteEscapedCharacter(TextWriter writer, char value) {
                switch (value) {
                    case '&':
                        writer.Write("&amp;");
                        break;
                    case '<':
                        writer.Write("&lt;");
                        break;
                    case '>':
                        writer.Write("&gt;");
                        break;
                    case '"':
                        writer.Write("&quot;");
                        break;
                    case '\'':
                        writer.Write("&apos;");
                        break;
                }
            }

        }

    }
}
