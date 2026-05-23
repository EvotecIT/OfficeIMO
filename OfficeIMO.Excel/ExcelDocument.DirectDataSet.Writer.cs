using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static class DirectDataSetWorkbookWriter {
            private const int XmlWriterBufferSize = 65536;
            private const string DateStyleAttribute = " s=\"1\"";
            private const string TimeStyleAttribute = " s=\"2\"";
            private const string CellValueDateStyleAttribute = " s=\"3\"";
            private const string CellValueTimeStyleAttribute = " s=\"4\"";
            private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            private static readonly DateTime ExcelMinimumSupportedDateTimeOffset = DateTime.FromOADate(2);

            internal static void Write(Stream stream, DirectDataSetWorkbookModel model, CancellationToken ct) {
                DirectColumnWritePlan[] columnWritePlans = CreateColumnWritePlans(model, ct);
                var sharedStrings = DirectSharedStringTable.Create(model, columnWritePlans, ct);
                using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
                WriteContentTypes(archive, model.Sheets, sharedStrings != null);
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
                WriteWorkbookRelationships(archive, model.Sheets.Count, sharedStrings != null);
                WriteStyles(archive);
                if (sharedStrings != null) {
                    WriteSharedStrings(archive, sharedStrings);
                }

                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < model.Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = model.Sheets[i];
                    WriteWorksheet(archive, sheet, model.DateTimeOffsetWriteStrategy, sharedStrings, columnWritePlans[i], ct);
                    if (sheet.HasTable) {
                        string sheetIndexText = InvariantNumberText.Get(sheet.Index);
                        WriteTextEntry(archive, "xl/worksheets/_rels/sheet" + sheetIndexText + ".xml.rels",
                            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table" + sheetIndexText + ".xml\"/>" +
                            "</Relationships>");
                        WriteTable(archive, sheet);
                    }
                }
            }

            private static void WriteContentTypes(ZipArchive archive, IReadOnlyList<DirectDataSetSheetModel> sheets, bool includeSharedStrings) {
                var builder = new StringBuilder(1024 + sheets.Count * 260);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
                builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                builder.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
                builder.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
                builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                if (includeSharedStrings) {
                    builder.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
                }

                foreach (var sheet in sheets) {
                    string sheetIndexText = InvariantNumberText.Get(sheet.Index);
                    builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                    builder.Append(sheetIndexText);
                    builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                    if (sheet.HasTable) {
                        builder.Append("<Override PartName=\"/xl/tables/table");
                        builder.Append(sheetIndexText);
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
                    string sheetIndexText = InvariantNumberText.Get(sheet.Index);
                    builder.Append("\" sheetId=\"");
                    builder.Append(sheetIndexText);
                    builder.Append("\" r:id=\"rId");
                    builder.Append(sheetIndexText);
                    builder.Append("\"/>");
                }

                builder.Append("</sheets></workbook>");
                WriteTextEntry(archive, "xl/workbook.xml", builder.ToString());
            }

            private static void WriteWorkbookRelationships(ZipArchive archive, int sheetCount, bool includeSharedStrings) {
                var builder = new StringBuilder(384 + sheetCount * 160);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 1; i <= sheetCount; i++) {
                    string indexText = InvariantNumberText.Get(i);
                    builder.Append("<Relationship Id=\"rId");
                    builder.Append(indexText);
                    builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
                    builder.Append(indexText);
                    builder.Append(".xml\"/>");
                }

                builder.Append("<Relationship Id=\"rId");
                builder.Append(InvariantNumberText.Get(sheetCount + 1));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                if (includeSharedStrings) {
                    builder.Append("<Relationship Id=\"rId");
                    builder.Append(InvariantNumberText.Get(sheetCount + 2));
                    builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
                }

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
                    "<cellXfs count=\"5\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"46\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>" +
                    "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                    "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>" +
                    "</styleSheet>");
            }

            private static void WriteSharedStrings(ZipArchive archive, DirectSharedStringTable sharedStrings) {
                var entry = archive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
                WriteInvariant(writer, sharedStrings.TotalStringReferences);
                writer.Write("\" uniqueCount=\"");
                WriteInvariant(writer, sharedStrings.Values.Length);
                writer.Write("\">");
                string[] values = sharedStrings.Values;
                for (int i = 0; i < values.Length; i++) {
                    string value = values[i];
                    writer.Write("<si><t");
                    if (NeedsPreserveSpace(value)) {
                        writer.Write(" xml:space=\"preserve\"");
                    }

                    writer.Write(">");
                    WriteSanitizedEscaped(writer, value);
                    writer.Write("</t></si>");
                }

                writer.Write("</sst>");
            }

            private static void WriteWorksheet(ZipArchive archive, DirectDataSetSheetModel sheet, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings, DirectColumnWritePlan columnWritePlan, CancellationToken ct) {
                var entry = archive.CreateEntry("xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml", CompressionLevel.Fastest);
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
                string?[]? styleAttributes = columnWritePlan.StyleAttributes;
                bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                DirectCellValueKind[] cellValueKinds = columnWritePlan.CellValueKinds;
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = 1;
                if (sheet.IncludeHeaders) {
                    const string headerRowReference = "1";
                    writer.Write("<row r=\"1\">");
                    for (int c = 0; c < columnCount; c++) {
                        WriteCell(writer, headerRowReference, cellReferencePrefixes[c], sheet.Table.GetColumnName(c), null, dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }

                bool hasBufferedRows = sheet.Table.TryGetBufferedRows(out DirectBufferedRows bufferedRows);
                if (hasBufferedRows
                    && !sheet.OmitBlankCells
                    && styleAttributes == null
                    && valueStyleColumns == null
                    && AllColumnsUseCellValueKind(cellValueKinds, DirectCellValueKind.String)) {
                    WritePlainStringBufferedRows(writer, bufferedRows, rowCount, columnCount, rowIndex, cellReferencePrefixes, sharedStrings, ct);
                    rowIndex += rowCount;
                } else if (hasBufferedRows
                    && !sheet.OmitBlankCells
                    && valueStyleColumns == null) {
                    WriteFixedKindBufferedRows(writer, bufferedRows, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    rowIndex += rowCount;
                } else if (valueStyleColumns == null) {
                    if (hasBufferedRows) {
                        if (sheet.OmitBlankCells) {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                bool rowStarted = false;
                                string rowReference = InvariantNumberText.Get(rowIndex);
                                object?[] bufferedRow = bufferedRows[sourceRowIndex];
                                for (int c = 0; c < columnCount; c++) {
                                    object? value = bufferedRow[c];
                                    if (IsBlankCellValue(value)) {
                                        continue;
                                    }

                                    if (!rowStarted) {
                                        writer.Write("<row r=\"");
                                        writer.Write(rowReference);
                                        writer.Write("\">");
                                        rowStarted = true;
                                    }

                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                                }

                                if (rowStarted) {
                                    writer.Write("</row>");
                                }

                                rowIndex++;
                            }
                        } else {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                string rowReference = InvariantNumberText.Get(rowIndex);
                                object?[] bufferedRow = bufferedRows[sourceRowIndex];
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                for (int c = 0; c < columnCount; c++) {
                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], bufferedRow[c], styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                                }

                                writer.Write("</row>");
                                rowIndex++;
                            }
                        }
                    } else if (sheet.Table.HasSourceRows) {
                        if (sheet.OmitBlankCells) {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                bool rowStarted = false;
                                string rowReference = InvariantNumberText.Get(rowIndex);
                                DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                                for (int c = 0; c < columnCount; c++) {
                                    object? value = sourceRow[c];
                                    if (IsBlankCellValue(value)) {
                                        continue;
                                    }

                                    if (!rowStarted) {
                                        writer.Write("<row r=\"");
                                        writer.Write(rowReference);
                                        writer.Write("\">");
                                        rowStarted = true;
                                    }

                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                                }

                                if (rowStarted) {
                                    writer.Write("</row>");
                                }

                                rowIndex++;
                            }
                        } else {
                            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                string rowReference = InvariantNumberText.Get(rowIndex);
                                DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                for (int c = 0; c < columnCount; c++) {
                                    WriteCell(writer, rowReference, cellReferencePrefixes[c], sourceRow[c], styleAttributes?[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                                }

                                writer.Write("</row>");
                                rowIndex++;
                            }
                        }
                    } else {
                        WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns: null, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                        rowIndex += rowCount;
                    }
                } else if (hasBufferedRows) {
                    if (sheet.OmitBlankCells) {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            bool rowStarted = false;
                            string rowReference = InvariantNumberText.Get(rowIndex);
                            object?[] bufferedRow = bufferedRows[sourceRowIndex];
                            for (int c = 0; c < columnCount; c++) {
                                object? value = bufferedRow[c];
                                if (IsBlankCellValue(value)) {
                                    continue;
                                }

                                if (!rowStarted) {
                                    writer.Write("<row r=\"");
                                    writer.Write(rowReference);
                                    writer.Write("\">");
                                    rowStarted = true;
                                }

                                WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                            }

                            if (rowStarted) {
                                writer.Write("</row>");
                            }

                            rowIndex++;
                        }
                    } else {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            string rowReference = InvariantNumberText.Get(rowIndex);
                            object?[] bufferedRow = bufferedRows[sourceRowIndex];
                            writer.Write("<row r=\"");
                            writer.Write(rowReference);
                            writer.Write("\">");
                            for (int c = 0; c < columnCount; c++) {
                                WriteCell(writer, rowReference, cellReferencePrefixes[c], bufferedRow[c], styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                            }

                            writer.Write("</row>");
                            rowIndex++;
                        }
                    }
                } else if (sheet.Table.HasSourceRows) {
                    if (sheet.OmitBlankCells) {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            bool rowStarted = false;
                            string rowReference = InvariantNumberText.Get(rowIndex);
                            DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                            for (int c = 0; c < columnCount; c++) {
                                object? value = sourceRow[c];
                                if (IsBlankCellValue(value)) {
                                    continue;
                                }

                                if (!rowStarted) {
                                    writer.Write("<row r=\"");
                                    writer.Write(rowReference);
                                    writer.Write("\">");
                                    rowStarted = true;
                                }

                                WriteCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                            }

                            if (rowStarted) {
                                writer.Write("</row>");
                            }

                            rowIndex++;
                        }
                    } else {
                        for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                            if (canCancel) {
                                ct.ThrowIfCancellationRequested();
                            }

                            string rowReference = InvariantNumberText.Get(rowIndex);
                            DataRow sourceRow = sheet.Table.GetSourceRow(sourceRowIndex)!;
                            writer.Write("<row r=\"");
                            writer.Write(rowReference);
                            writer.Write("\">");
                            for (int c = 0; c < columnCount; c++) {
                                WriteCell(writer, rowReference, cellReferencePrefixes[c], sourceRow[c], styleAttributes?[c], valueStyleColumns[c], cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                            }

                            writer.Write("</row>");
                            rowIndex++;
                        }
                    }
                } else {
                    WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    rowIndex += rowCount;
                }

                writer.Write("</sheetData>");
                if (sheet.HasTable) {
                    writer.Write("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>");
                }

                writer.Write("</worksheet>");
            }

            private static void WriteFixedKindBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (styleAttributes == null) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        object?[] bufferedRow = bufferedRows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            WriteCellValue(writer, bufferedRow[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    object?[] bufferedRow = bufferedRows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        string? styleAttribute = styleAttributes[c];
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        WriteCellValue(writer, bufferedRow[c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WritePlainStringBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    object?[] bufferedRow = bufferedRows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        object? value = bufferedRow[c];
                        if (value is string text) {
                            WriteStringCellValue(writer, text, sharedStrings);
                        } else if (value == null || value == DBNull.Value) {
                            writer.Write(" t=\"str\"><v/></c>");
                        } else {
                            WriteStringCell(writer, value.ToString() ?? string.Empty, validateLength: true);
                        }
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static bool AllColumnsUseCellValueKind(DirectCellValueKind[] cellValueKinds, DirectCellValueKind expected) {
                for (int i = 0; i < cellValueKinds.Length; i++) {
                    if (cellValueKinds[i] != expected) {
                        return false;
                    }
                }

                return true;
            }

            private static bool IsBlankCellValue(object? value) {
                return value == null || value == DBNull.Value;
            }

            private static string[] CreateCellReferencePrefixes(int columnCount) {
                var columns = new string[columnCount];
                for (int i = 0; i < columnCount; i++) {
                    columns[i] = "<c r=\"" + A1.ColumnIndexToLetters(i + 1);
                }

                return columns;
            }

            private static DirectColumnWritePlan CreateColumnWritePlan(DirectDataSetTableModel table, bool useCellValueNumberFormats) {
                int columnCount = table.ColumnCount;
                var kinds = new DirectCellValueKind[columnCount];
                string?[]? styleAttributes = null;
                bool[]? valueStyleColumns = null;

                for (int i = 0; i < columnCount; i++) {
                    Type dataType = table.GetColumnType(i);
                    DirectCellValueKind cellValueKind = GetCellValueKind(dataType);
                    bool useValueStyle = dataType == typeof(object);
                    if (useValueStyle && TryInferObjectColumnCellValueKind(table, i, out DirectCellValueKind inferredKind)) {
                        cellValueKind = inferredKind;
                        useValueStyle = false;
                    }

                    kinds[i] = cellValueKind;

                    string? styleAttribute = GetStyleAttribute(cellValueKind, useCellValueNumberFormats);
                    if (styleAttribute != null) {
                        styleAttributes ??= new string?[columnCount];
                        styleAttributes[i] = styleAttribute;
                    }

                    if (useValueStyle) {
                        valueStyleColumns ??= new bool[columnCount];
                        valueStyleColumns[i] = true;
                    }
                }

                return new DirectColumnWritePlan(kinds, styleAttributes, valueStyleColumns);
            }

            private static bool TryInferObjectColumnCellValueKind(DirectDataSetTableModel table, int columnIndex, out DirectCellValueKind inferredKind) {
                inferredKind = DirectCellValueKind.Object;
                int rowCount = table.RowCount;
                bool sawValue = false;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    object? value = table.GetValue(rowIndex, columnIndex);
                    if (value == null || value == DBNull.Value) {
                        continue;
                    }

                    DirectCellValueKind valueKind = GetCellValueKind(value.GetType());
                    if (valueKind == DirectCellValueKind.Object) {
                        return false;
                    }

                    if (!sawValue) {
                        inferredKind = valueKind;
                        sawValue = true;
                        continue;
                    }

                    if (inferredKind != valueKind) {
                        return false;
                    }
                }

                return sawValue;
            }

            private static DirectCellValueKind GetCellValueKind(Type dataType) {
                if (dataType == typeof(string)) return DirectCellValueKind.String;
                if (dataType == typeof(bool)) return DirectCellValueKind.Boolean;
                if (dataType == typeof(DateTime)) return DirectCellValueKind.DateTime;
                if (dataType == typeof(DateTimeOffset)) return DirectCellValueKind.DateTimeOffset;
                if (dataType == typeof(TimeSpan)) return DirectCellValueKind.TimeSpan;
                if (dataType == typeof(double)) return DirectCellValueKind.Double;
                if (dataType == typeof(float)) return DirectCellValueKind.Float;
                if (dataType == typeof(decimal)) return DirectCellValueKind.Decimal;
                if (dataType == typeof(sbyte)) return DirectCellValueKind.SByte;
                if (dataType == typeof(byte)) return DirectCellValueKind.Byte;
                if (dataType == typeof(short)) return DirectCellValueKind.Int16;
                if (dataType == typeof(ushort)) return DirectCellValueKind.UInt16;
                if (dataType == typeof(int)) return DirectCellValueKind.Int32;
                if (dataType == typeof(uint)) return DirectCellValueKind.UInt32;
                if (dataType == typeof(long)) return DirectCellValueKind.Int64;
                if (dataType == typeof(ulong)) return DirectCellValueKind.UInt64;
#if NET6_0_OR_GREATER
                if (dataType == typeof(DateOnly)) return DirectCellValueKind.DateOnly;
                if (dataType == typeof(TimeOnly)) return DirectCellValueKind.TimeOnly;
#endif
                return DirectCellValueKind.Object;
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
                var entry = archive.CreateEntry("xl/tables/table" + InvariantNumberText.Get(sheet.Index) + ".xml", CompressionLevel.Fastest);
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

            private enum DirectCellValueKind {
                Object,
                String,
                Boolean,
                DateTime,
                DateTimeOffset,
                TimeSpan,
                Double,
                Float,
                Decimal,
                SByte,
                Byte,
                Int16,
                UInt16,
                Int32,
                UInt32,
                Int64,
                UInt64,
#if NET6_0_OR_GREATER
                DateOnly,
                TimeOnly
#endif
            }

            private readonly struct DirectColumnWritePlan {
                internal DirectColumnWritePlan(DirectCellValueKind[] cellValueKinds, string?[]? styleAttributes, bool[]? valueStyleColumns) {
                    CellValueKinds = cellValueKinds;
                    StyleAttributes = styleAttributes;
                    ValueStyleColumns = valueStyleColumns;
                }

                internal DirectCellValueKind[] CellValueKinds { get; }

                internal string?[]? StyleAttributes { get; }

                internal bool[]? ValueStyleColumns { get; }
            }

            private sealed class DirectSharedStringTable {
                private const int MinimumStringReferences = 512;
                private const int MinimumDuplicateReferences = 128;
                private const int MinimumDuplicateCharacters = 4096;
                private const int MaximumSeenOnceCandidates = 8192;
                private const int MinimumEarlyUniqueHeavyStringReferences = 16384;
                private const long MinimumDuplicateCharacterShareNumerator = 3L;
                private const long MinimumDuplicateCharacterShareDenominator = 5L;
                private readonly Dictionary<string, int> _indexes;

                private DirectSharedStringTable(Dictionary<string, int> indexes, string[] values, int totalStringReferences) {
                    _indexes = indexes;
                    Values = values;
                    TotalStringReferences = totalStringReferences;
                }

                internal string[] Values { get; }

                internal int TotalStringReferences { get; }

                internal bool TryGetIndex(string value, out int index) => _indexes.TryGetValue(value, out index);

                internal static DirectSharedStringTable? Create(DirectDataSetWorkbookModel model, IReadOnlyList<DirectColumnWritePlan> columnWritePlans, CancellationToken ct) {
                    if (!CanReachMinimumStringReferences(model, columnWritePlans)) {
                        return null;
                    }

                    var seenOnce = new HashSet<string>(StringComparer.Ordinal);
                    var sharedCounts = new Dictionary<string, int>(StringComparer.Ordinal);
                    int totalStringReferences = 0;
                    int duplicateReferences = 0;
                    long totalStringCharacters = 0L;
                    long duplicateCharacters = 0L;
                    bool canCancel = ct.CanBeCanceled;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var sheet = model.Sheets[sheetIndex];
                        int columnCount = sheet.Table.ColumnCount;
                        if (sheet.IncludeHeaders) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                NoteString(sheet.Table.GetColumnName(columnIndex), forceShared: true);
                            }
                        }

                        int[]? stringColumnIndexes = CreateSharedStringCandidateColumnIndexes(columnWritePlans[sheetIndex]);
                        if (stringColumnIndexes == null) {
                            continue;
                        }

                        int rowCount = sheet.Table.RowCount;
                        int stringColumnCount = stringColumnIndexes.Length;
                        bool hasBufferedRows = sheet.Table.TryGetBufferedRows(out DirectBufferedRows bufferedRows);
                        if (hasBufferedRows) {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                object?[] bufferedRow = bufferedRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (bufferedRow[columnIndex] is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.HasSourceRows) {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                DataRow sourceRow = sheet.Table.GetSourceRow(rowIndex)!;
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (sourceRow[columnIndex] is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.TryGetExactDictionaryRows(out var exactDictionaryRows)) {
                            string[] columnNames = sheet.Table.CreateColumnNameArray();
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                Dictionary<string, object?> row = exactDictionaryRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (row.TryGetValue(columnNames[columnIndex], out object? value)
                                        && value is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.TryGetDictionaryRows(out var dictionaryRows)) {
                            string[] columnNames = sheet.Table.CreateColumnNameArray();
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                IReadOnlyDictionary<string, object?> row = dictionaryRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (row.TryGetValue(columnNames[columnIndex], out object? value)
                                        && value is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else if (sheet.Table.TryGetLegacyDictionaryRows(out var legacyDictionaryRows)) {
                            string[] columnNames = sheet.Table.CreateColumnNameArray();
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                System.Collections.IDictionary row = legacyDictionaryRows[rowIndex];
                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (DirectDataSetTableModel.GetLegacyDictionaryValue(row, columnNames[columnIndex]) is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        } else {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                                if (canCancel) {
                                    ct.ThrowIfCancellationRequested();
                                }

                                for (int stringColumnIndex = 0; stringColumnIndex < stringColumnCount; stringColumnIndex++) {
                                    int columnIndex = stringColumnIndexes[stringColumnIndex];
                                    if (sheet.Table.GetValue(rowIndex, columnIndex) is string text) {
                                        NoteString(text);
                                    }
                                }

                                if (ShouldAbandonSharedStrings()) {
                                    return null;
                                }
                            }
                        }
                    }

                    if (totalStringReferences < MinimumStringReferences || sharedCounts.Count == 0) {
                        return null;
                    }

                    if (duplicateReferences < MinimumDuplicateReferences && duplicateCharacters < MinimumDuplicateCharacters) {
                        return null;
                    }

                    if (duplicateCharacters * MinimumDuplicateCharacterShareDenominator < totalStringCharacters * MinimumDuplicateCharacterShareNumerator) {
                        return null;
                    }

                    var indexes = new Dictionary<string, int>(sharedCounts.Count, StringComparer.Ordinal);
                    var values = new string[sharedCounts.Count];
                    int nextIndex = 0;
                    int sharedStringReferences = 0;
                    foreach (var entry in sharedCounts) {
                        CoerceValueHelper.ValidateSharedStringLength(entry.Key, "value");
                        values[nextIndex] = entry.Key;
                        indexes[entry.Key] = nextIndex;
                        sharedStringReferences += entry.Value;
                        nextIndex++;
                    }

                    return new DirectSharedStringTable(indexes, values, sharedStringReferences);

                    void NoteString(string text, bool forceShared = false) {
                        totalStringReferences++;
                        totalStringCharacters += text.Length;
                        if (sharedCounts.TryGetValue(text, out int count)) {
                            sharedCounts[text] = count + 1;
                            duplicateReferences++;
                            duplicateCharacters += text.Length;
                            return;
                        }

                        if (forceShared) {
                            if (seenOnce.Remove(text)) {
                                sharedCounts.Add(text, 2);
                                duplicateReferences++;
                                duplicateCharacters += text.Length;
                            } else {
                                sharedCounts.Add(text, 1);
                            }

                            return;
                        }

                        if (seenOnce.Count >= MaximumSeenOnceCandidates) {
                            if (seenOnce.Remove(text)) {
                                sharedCounts.Add(text, 2);
                                duplicateReferences++;
                                duplicateCharacters += text.Length;
                            }

                            return;
                        }

                        if (!seenOnce.Add(text)) {
                            seenOnce.Remove(text);
                            sharedCounts.Add(text, 2);
                            duplicateReferences++;
                            duplicateCharacters += text.Length;
                        }
                    }

                    bool ShouldAbandonSharedStrings() {
                        return totalStringReferences >= MinimumEarlyUniqueHeavyStringReferences
                            && seenOnce.Count >= MaximumSeenOnceCandidates
                            && duplicateCharacters * MinimumDuplicateCharacterShareDenominator < totalStringCharacters * MinimumDuplicateCharacterShareNumerator;
                    }
                }

                private static bool CanReachMinimumStringReferences(DirectDataSetWorkbookModel model, IReadOnlyList<DirectColumnWritePlan> columnWritePlans) {
                    long possibleStringReferences = 0L;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var sheet = model.Sheets[sheetIndex];
                        int columnCount = sheet.Table.ColumnCount;
                        if (sheet.IncludeHeaders) {
                            possibleStringReferences += columnCount;
                        }

                        int[]? stringColumnIndexes = CreateSharedStringCandidateColumnIndexes(columnWritePlans[sheetIndex]);
                        if (stringColumnIndexes != null) {
                            int rowCount = sheet.Table.RowCount;
                            possibleStringReferences += (long)rowCount * stringColumnIndexes.Length;
                        }

                        if (possibleStringReferences >= MinimumStringReferences) {
                            return true;
                        }
                    }

                    return false;
                }

                private static int[]? CreateSharedStringCandidateColumnIndexes(DirectColumnWritePlan columnWritePlan) {
                    DirectCellValueKind[] kinds = columnWritePlan.CellValueKinds;
                    bool[]? valueStyleColumns = columnWritePlan.ValueStyleColumns;
                    int[]? indexes = null;
                    int count = 0;
                    for (int i = 0; i < kinds.Length; i++) {
                        bool canContainString = kinds[i] == DirectCellValueKind.String
                            || valueStyleColumns?[i] == true;
                        if (!canContainString) {
                            continue;
                        }

                        indexes ??= new int[kinds.Length];
                        indexes[count++] = i;
                    }

                    if (indexes == null) {
                        return null;
                    }

                    if (count == indexes.Length) {
                        return indexes;
                    }

                    Array.Resize(ref indexes, count);
                    return indexes;
                }
            }

            private static string? GetStyleAttribute(DirectCellValueKind cellValueKind, bool useCellValueNumberFormats) {
                switch (cellValueKind) {
                    case DirectCellValueKind.DateTime:
                    case DirectCellValueKind.DateTimeOffset:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case DirectCellValueKind.TimeSpan:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#if NET6_0_OR_GREATER
                    case DirectCellValueKind.DateOnly:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case DirectCellValueKind.TimeOnly:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#endif
                    default:
                        return null;
                }
            }

            private static DirectColumnWritePlan[] CreateColumnWritePlans(DirectDataSetWorkbookModel model, CancellationToken ct) {
                var plans = new DirectColumnWritePlan[model.Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < plans.Length; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    DirectDataSetSheetModel sheet = model.Sheets[i];
                    plans[i] = CreateColumnWritePlan(sheet.Table, sheet.UseCellValueNumberFormats);
                }

                return plans;
            }

            private static void WriteDirectValueRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                if (sheet.Table.TryGetExactDictionaryRows(out var exactDictionaryRows)) {
                    WriteExactDictionaryValueRows(writer, exactDictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetDictionaryRows(out var dictionaryRows)) {
                    WriteDictionaryValueRows(writer, dictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetLegacyDictionaryRows(out var legacyDictionaryRows)) {
                    WriteLegacyDictionaryValueRows(writer, legacyDictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        for (int c = 0; c < columnCount; c++) {
                            object? value = sheet.Table.GetValue(sourceRowIndex, c);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = sheet.Table.GetValue(sourceRowIndex, c);
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteExactDictionaryValueRows(
                TextWriter writer,
                IReadOnlyList<Dictionary<string, object?>> rows,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindExactDictionaryRows(writer, rows, rowCount, columnCount, rowIndex, cellReferencePrefixes, columnNames, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        Dictionary<string, object?> row = rows[sourceRowIndex];
                        for (int c = 0; c < columnCount; c++) {
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    Dictionary<string, object?> row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteFixedKindExactDictionaryRows(
                TextWriter writer,
                IReadOnlyList<Dictionary<string, object?>> rows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (styleAttributes == null) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        Dictionary<string, object?> row = rows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    Dictionary<string, object?> row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        string? styleAttribute = styleAttributes[c];
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDictionaryValueRows(
                TextWriter writer,
                IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindDictionaryRows(writer, rows, rowCount, columnCount, rowIndex, cellReferencePrefixes, columnNames, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                        for (int c = 0; c < columnCount; c++) {
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteFixedKindDictionaryRows(
                TextWriter writer,
                IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (styleAttributes == null) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                                ? dictionaryValue
                                : null;
                            WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    IReadOnlyDictionary<string, object?> row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        string? styleAttribute = styleAttributes[c];
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        object? value = row.TryGetValue(columnNames[c], out object? dictionaryValue)
                            ? dictionaryValue
                            : null;
                        WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteLegacyDictionaryValueRows(
                TextWriter writer,
                IReadOnlyList<System.Collections.IDictionary> rows,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindLegacyDictionaryRows(writer, rows, rowCount, columnCount, rowIndex, cellReferencePrefixes, columnNames, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.OmitBlankCells) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        bool rowStarted = false;
                        string rowReference = InvariantNumberText.Get(rowIndex);
                        System.Collections.IDictionary row = rows[sourceRowIndex];
                        for (int c = 0; c < columnCount; c++) {
                            object? value = DirectDataSetTableModel.GetLegacyDictionaryValue(row, columnNames[c]);
                            if (IsBlankCellValue(value)) {
                                continue;
                            }

                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                                rowStarted = true;
                            }

                            WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        if (rowStarted) {
                            writer.Write("</row>");
                        }

                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    System.Collections.IDictionary row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        object? value = DirectDataSetTableModel.GetLegacyDictionaryValue(row, columnNames[c]);
                        WriteDirectValueCell(writer, rowReference, cellReferencePrefixes[c], value, styleAttributes?[c], valueStyleColumns?[c] ?? false, cellValueKinds[c], sheet.UseCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteFixedKindLegacyDictionaryRows(
                TextWriter writer,
                IReadOnlyList<System.Collections.IDictionary> rows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string[] columnNames,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (styleAttributes == null) {
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        System.Collections.IDictionary row = rows[sourceRowIndex];
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = DirectDataSetTableModel.GetLegacyDictionaryValue(row, columnNames[c]);
                            WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                    }

                    return;
                }

                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
                    System.Collections.IDictionary row = rows[sourceRowIndex];
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    for (int c = 0; c < columnCount; c++) {
                        writer.Write(cellReferencePrefixes[c]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        string? styleAttribute = styleAttributes[c];
                        if (styleAttribute != null) {
                            writer.Write(styleAttribute);
                        }

                        object? value = DirectDataSetTableModel.GetLegacyDictionaryValue(row, columnNames[c]);
                        WriteCellValue(writer, value, cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                }
            }

            private static void WriteDirectValueCell(
                TextWriter writer,
                string rowReference,
                string cellReferencePrefix,
                object? value,
                string? styleAttribute,
                bool valueStyleColumn,
                DirectCellValueKind cellValueKind,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings) {
                if (valueStyleColumn) {
                    WriteCell(writer, rowReference, cellReferencePrefix, value, styleAttribute, valueStyleColumn, cellValueKind, useCellValueNumberFormats, dateTimeOffsetWriteStrategy, sharedStrings);
                } else {
                    WriteCell(writer, rowReference, cellReferencePrefix, value, styleAttribute, cellValueKind, dateTimeOffsetWriteStrategy, sharedStrings);
                }
            }

            private static string? CreateStyleAttributeForValue(object? value, bool useCellValueNumberFormats) {
                switch (value) {
                    case DateTime:
                    case DateTimeOffset:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case TimeSpan:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#if NET6_0_OR_GREATER
                    case DateOnly:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case TimeOnly:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#endif
                    default:
                        return null;
                }
            }

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                if (styleAttribute != null) {
                    writer.Write(styleAttribute);
                }

                WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, sharedStrings);
            }

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, DirectCellValueKind cellValueKind, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                if (styleAttribute != null) {
                    writer.Write(styleAttribute);
                }

                WriteCellValue(writer, value, cellValueKind, dateTimeOffsetWriteStrategy, sharedStrings);
            }

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, bool useValueStyle, DirectCellValueKind cellValueKind, bool useCellValueNumberFormats, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                string? effectiveStyleAttribute = styleAttribute ?? (useValueStyle ? CreateStyleAttributeForValue(value, useCellValueNumberFormats) : null);
                if (effectiveStyleAttribute != null) {
                    writer.Write(effectiveStyleAttribute);
                }

                if (useValueStyle) {
                    WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, sharedStrings);
                } else {
                    WriteCellValue(writer, value, cellValueKind, dateTimeOffsetWriteStrategy, sharedStrings);
                }
            }

            private static void WriteCellValue(TextWriter writer, object? value, DirectCellValueKind cellValueKind, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings) {
                if (value == null || value == DBNull.Value) {
                    writer.Write(" t=\"str\"><v/></c>");
                    return;
                }

                switch (cellValueKind) {
                    case DirectCellValueKind.String:
                        if (value is string stringValue) {
                            WriteStringCellValue(writer, stringValue, sharedStrings);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Boolean:
                        if (value is bool boolValue) {
                            writer.Write(boolValue ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                            return;
                        }

                        break;
                    case DirectCellValueKind.DateTime:
                        if (value is DateTime dateTime) {
                            WriteRawValueCell(writer, dateTime.ToOADate());
                            return;
                        }

                        break;
                    case DirectCellValueKind.DateTimeOffset:
                        if (value is DateTimeOffset dateTimeOffset) {
                            WriteDateTimeOffsetCellValue(writer, dateTimeOffset, dateTimeOffsetWriteStrategy);
                            return;
                        }

                        break;
                    case DirectCellValueKind.TimeSpan:
                        if (value is TimeSpan timeSpan) {
                            WriteRawValueCell(writer, timeSpan.TotalDays);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Double:
                        if (value is double doubleValue) {
                            WriteRawValueCell(writer, doubleValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Float:
                        if (value is float floatValue) {
                            WriteRawValueCell(writer, floatValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Decimal:
                        if (value is decimal decimalValue) {
                            WriteRawValueCell(writer, decimalValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.SByte:
                        if (value is sbyte sbyteValue) {
                            WriteRawValueCell(writer, sbyteValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Byte:
                        if (value is byte byteValue) {
                            WriteRawValueCell(writer, byteValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Int16:
                        if (value is short shortValue) {
                            WriteRawValueCell(writer, shortValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.UInt16:
                        if (value is ushort ushortValue) {
                            WriteRawValueCell(writer, ushortValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Int32:
                        if (value is int intValue) {
                            WriteRawValueCell(writer, intValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.UInt32:
                        if (value is uint uintValue) {
                            WriteRawValueCell(writer, uintValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Int64:
                        if (value is long longValue) {
                            WriteRawValueCell(writer, longValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.UInt64:
                        if (value is ulong ulongValue) {
                            WriteRawValueCell(writer, ulongValue);
                            return;
                        }

                        break;
#if NET6_0_OR_GREATER
                    case DirectCellValueKind.DateOnly:
                        if (value is DateOnly dateOnly) {
                            WriteRawValueCell(writer, dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate());
                            return;
                        }

                        break;
                    case DirectCellValueKind.TimeOnly:
                        if (value is TimeOnly timeOnly) {
                            WriteRawValueCell(writer, timeOnly.ToTimeSpan().TotalDays);
                            return;
                        }

                        break;
#endif
                }

                WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, sharedStrings);
            }

            private static void WriteCellValue(TextWriter writer, object? value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings) {
                switch (value) {
                    case null:
                    case DBNull:
                        writer.Write(" t=\"str\"><v/></c>");
                        return;
                    case string stringValue:
                        WriteStringCellValue(writer, stringValue, sharedStrings);
                        return;
                    case bool boolValue:
                        writer.Write(boolValue ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                        return;
                    case DateTime dateTime:
                        WriteRawValueCell(writer, dateTime.ToOADate());
                        return;
                    case DateTimeOffset dateTimeOffset:
                        WriteDateTimeOffsetCellValue(writer, dateTimeOffset, dateTimeOffsetWriteStrategy);
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
                    case long longValue:
                        WriteRawValueCell(writer, longValue);
                        return;
                    case ulong ulongValue:
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
                        WriteStringCell(writer, value.ToString() ?? string.Empty, validateLength: true);
                        return;
                }
            }

            private static void WriteStringCellValue(TextWriter writer, string value, DirectSharedStringTable? sharedStrings) {
                if (sharedStrings != null && sharedStrings.TryGetIndex(value, out int sharedStringIndex)) {
                    WriteSharedStringCell(writer, sharedStringIndex);
                } else {
                    WriteStringCell(writer, value, validateLength: true);
                }
            }

            private static void WriteDateTimeOffsetCellValue(TextWriter writer, DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                if (!TryGetDateTimeOffsetSerial(value, dateTimeOffsetWriteStrategy, out double dateTimeOffsetSerial)) {
                    WriteStringCell(writer, value.ToString("o", CultureInfo.InvariantCulture), validateLength: true);
                    return;
                }

                WriteRawValueCell(writer, dateTimeOffsetSerial);
            }

            private static bool TryGetDateTimeOffsetSerial(DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, out double serial) {
                try {
                    if (value.UtcDateTime < ExcelMinimumSupportedDateTimeOffset) {
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

            private static void WriteStringCell(TextWriter writer, string text, bool validateLength) {
                if (validateLength) {
                    CoerceValueHelper.ValidateSharedStringLength(text, "value");
                }

                if (text.Length == 0) {
                    writer.Write(" t=\"str\"><v/></c>");
                    return;
                }

                writer.Write(" t=\"str\"><v>");
                WriteSanitizedEscaped(writer, text);
                writer.Write("</v></c>");
            }

            private static void WriteSharedStringCell(TextWriter writer, int sharedStringIndex) {
                writer.Write(" t=\"s\"><v>");
                WriteInvariant(writer, sharedStringIndex);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, double value) {
                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, float value) {
                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, decimal value) {
                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, int value) {
                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, long value) {
                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, ulong value) {
                writer.Write(" t=\"n\"><v>");
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
                if (InvariantNumberText.TryGet(value, out string text)) {
                    writer.Write(text);
                    return;
                }

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
                if (InvariantNumberText.TryGet(value, out string text)) {
                    writer.Write(text);
                    return;
                }

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
                if (InvariantNumberText.TryGet(value, out string text)) {
                    writer.Write(text);
                    return;
                }

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

            private static bool NeedsPreserveSpace(string value) {
                return value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]));
            }

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
