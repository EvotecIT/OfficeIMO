using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static class DirectDataSetWorkbookWriter {
            private const int XmlWriterBufferSize = 32768;
            private const int StackallocTextEntryByteLimit = 4096;
            private const string DateStyleAttribute = " s=\"1\"";
            private const string TimeStyleAttribute = " s=\"2\"";
            private const string CellValueDateStyleAttribute = " s=\"3\"";
            private const string CellValueTimeStyleAttribute = " s=\"4\"";
            private const int CachedRawIntegerCellLimit = 65536;
            private const int CachedSharedStringCellLimit = 1024;
            private const int CachedCellReferencePrefixColumnLimit = 64;
            private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            private static readonly DateTime ExcelMinimumSupportedDateTimeOffset = DateTime.FromOADate(2);
            private static readonly string[] RawNonNegativeIntegerCellCache = CreateRawNonNegativeIntegerCellCache();
            private static readonly string?[] SharedStringCellCache = new string?[CachedSharedStringCellLimit];
            private static readonly string[][] CellReferencePrefixCache = CreateCellReferencePrefixCache();

            internal static void Write(Stream stream, DirectDataSetWorkbookModel model, CancellationToken ct) {
                DirectStylePlan stylePlan = DirectStylePlan.Create(model);
                DirectColumnWritePlan[] columnWritePlans = CreateColumnWritePlans(model, stylePlan, ct);
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
                WriteStyles(archive, stylePlan);
                if (sharedStrings != null) {
                    WriteSharedStrings(archive, sharedStrings);
                }

                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < model.Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = model.Sheets[i];
                    WriteWorksheet(archive, sheet, model.DateTimeOffsetWriteStrategy, sharedStrings, stylePlan, columnWritePlans[i], ct);
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

            internal static bool TryCreateExtendedWritePlan(DirectDataSetWorkbookModel model, CancellationToken ct, out ExtendedDirectWritePlan? plan) {
                DirectStylePlan stylePlan = DirectStylePlan.Create(model);
                DirectColumnWritePlan[] columnWritePlans = CreateColumnWritePlans(model, stylePlan, ct);
                var sharedStrings = DirectSharedStringTable.Create(model, columnWritePlans, ct);

                plan = new ExtendedDirectWritePlan(model, stylePlan, columnWritePlans, sharedStrings);
                return true;
            }

            internal static void WriteExtendedStyles(ZipArchive archive, ExtendedDirectWritePlan plan) {
                WriteStyles(archive, plan.StylePlan);
            }

            internal static void WriteExtendedSharedStrings(ZipArchive archive, ExtendedDirectWritePlan plan) {
                if (plan.SharedStrings != null) {
                    WriteSharedStrings(archive, plan.SharedStrings);
                }
            }

            internal static void WriteExtendedWorksheet(ZipArchive archive, ExtendedDirectWritePlan plan, DirectDataSetSheetModel sheet, string worksheetPath, string? tableRelationshipId, CancellationToken ct) {
                int sheetIndex = -1;
                for (int i = 0; i < plan.Model.Sheets.Count; i++) {
                    if (ReferenceEquals(plan.Model.Sheets[i], sheet)) {
                        sheetIndex = i;
                        break;
                    }
                }

                if (sheetIndex < 0 || sheetIndex >= plan.ColumnWritePlans.Length) {
                    throw new InvalidOperationException("The direct worksheet is not part of the extended direct write plan.");
                }

                WriteWorksheet(archive, sheet, plan.Model.DateTimeOffsetWriteStrategy, plan.SharedStrings, plan.StylePlan, plan.ColumnWritePlans[sheetIndex], ct, worksheetPath, tableRelationshipId);
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

            private static void WriteStyles(ZipArchive archive, DirectStylePlan stylePlan) {
                var builder = new StringBuilder(1024 + stylePlan.CustomNumberFormats.Count * 160);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                builder.Append("<numFmts count=\"");
                builder.Append((2 + stylePlan.CustomNumberFormats.Count).ToString(CultureInfo.InvariantCulture));
                builder.Append("\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd hh:mm\"/><numFmt numFmtId=\"165\" formatCode=\"[h]:mm:ss\"/>");
                for (int i = 0; i < stylePlan.CustomNumberFormats.Count; i++) {
                    builder.Append("<numFmt numFmtId=\"");
                    builder.Append((166 + i).ToString(CultureInfo.InvariantCulture));
                    builder.Append("\" formatCode=\"");
                    AppendEscaped(builder, stylePlan.CustomNumberFormats[i]);
                    builder.Append("\"/>");
                }

                builder.Append("</numFmts>");
                builder.Append("<fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts>");
                builder.Append("<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>");
                builder.Append("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");
                builder.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
                builder.Append("<cellXfs count=\"");
                builder.Append((5 + stylePlan.CustomNumberFormats.Count).ToString(CultureInfo.InvariantCulture));
                builder.Append("\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"46\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
                for (int i = 0; i < stylePlan.CustomNumberFormats.Count; i++) {
                    builder.Append("<xf numFmtId=\"");
                    builder.Append((166 + i).ToString(CultureInfo.InvariantCulture));
                    builder.Append("\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
                }

                builder.Append("</cellXfs>");
                builder.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
                builder.Append("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");
                builder.Append("</styleSheet>");
                WriteTextEntry(archive, "xl/styles.xml", builder.ToString());
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

            private static void WriteWorksheet(ZipArchive archive, DirectDataSetSheetModel sheet, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, DirectSharedStringTable? sharedStrings, DirectStylePlan stylePlan, DirectColumnWritePlan columnWritePlan, CancellationToken ct, string? worksheetPath = null, string? tableRelationshipId = null) {
                var entry = archive.CreateEntry(worksheetPath ?? "xl/worksheets/sheet" + InvariantNumberText.Get(sheet.Index) + ".xml", CompressionLevel.Fastest);
                using var stream = entry.Open();
                using var writer = new StreamWriter(stream, Utf8NoBom, XmlWriterBufferSize);

                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                var metadata = sheet.Metadata;
                if (!string.IsNullOrEmpty(metadata?.SheetPropertiesXml)) {
                    writer.Write(metadata!.SheetPropertiesXml);
                }

                writer.Write("<dimension ref=\"");
                WriteEscaped(writer, GetWorksheetDimension(sheet));
                writer.Write("\"/>");
                if (!string.IsNullOrEmpty(metadata?.SheetViewsXml)) {
                    writer.Write(metadata!.SheetViewsXml);
                }

                if (!string.IsNullOrEmpty(metadata?.SheetFormatPropertiesXml)) {
                    writer.Write(metadata!.SheetFormatPropertiesXml);
                }

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
                var overlayCells = metadata?.OverlayCells;
                var overlayCellsByRow = CreateOverlayCellsByRow(overlayCells);
                int directLastRow = rowCount + (sheet.IncludeHeaders ? 1 : 0);
                bool hasInlineOverlayCells = HasOverlayCellsAtOrBefore(overlayCells, directLastRow);
                if (sheet.IncludeHeaders) {
                    const string headerRowReference = "1";
                    writer.Write("<row r=\"1\">");
                    for (int c = 0; c < columnCount; c++) {
                        WriteCell(writer, headerRowReference, cellReferencePrefixes[c], sheet.Table.GetColumnName(c), null, dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    WriteOverlayCellsForRow(writer, overlayCellsByRow, 1, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings);
                    writer.Write("</row>");
                    rowIndex++;
                }

                bool hasBufferedRows = sheet.Table.TryGetBufferedRows(out DirectBufferedRows bufferedRows);
                bool hasCellValueRows = sheet.Table.TryGetCellValueRows(out DirectCellValueRows cellValueRows);
                if (hasInlineOverlayCells) {
                    WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings, ct, overlayCellsByRow);
                    rowIndex += rowCount;
                } else if (hasBufferedRows
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
                } else if (hasCellValueRows
                    && !sheet.OmitBlankCells
                    && valueStyleColumns == null) {
                    WriteFixedKindCellValueRows(writer, cellValueRows, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, sharedStrings, ct);
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
                        WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns: null, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings, ct);
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
                    WriteDirectValueRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    rowIndex += rowCount;
                }

                WriteOverlayCells(writer, overlayCells, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings, directLastRow);
                writer.Write("</sheetData>");
                if (!string.IsNullOrEmpty(metadata?.AutoFilterXml)) {
                    writer.Write(metadata!.AutoFilterXml);
                }

                if (metadata != null) {
                    for (int i = 0; i < metadata.ConditionalFormattingXml.Count; i++) {
                        writer.Write(metadata.ConditionalFormattingXml[i]);
                    }
                }

                if (!string.IsNullOrEmpty(metadata?.DataValidationsXml)) {
                    writer.Write(metadata!.DataValidationsXml);
                }

                if (metadata != null) {
                    for (int i = 0; i < metadata.PostDataValidationXml.Count; i++) {
                        writer.Write(metadata.PostDataValidationXml[i]);
                    }
                }

                if (!string.IsNullOrEmpty(metadata?.DrawingXml)) {
                    writer.Write(metadata!.DrawingXml);
                }

                if (sheet.HasTable) {
                    writer.Write("<tableParts count=\"1\"><tablePart r:id=\"");
                    WriteEscaped(writer, tableRelationshipId ?? "rId1");
                    writer.Write("\"/></tableParts>");
                }

                writer.Write("</worksheet>");
            }

            private static void WriteFixedKindCellValueRows(
                TextWriter writer,
                DirectCellValueRows cellValueRows,
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
                    object?[] values = cellValueRows.Values;
                    int rowOffset = 0;
                    for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        string rowReference = InvariantNumberText.Get(rowIndex);
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            WriteCellValue(writer, values[rowOffset + c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                        }

                        writer.Write("</row>");
                        rowIndex++;
                        rowOffset += columnCount;
                    }

                    return;
                }

                object?[] styledValues = cellValueRows.Values;
                int styledRowOffset = 0;
                for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    string rowReference = InvariantNumberText.Get(rowIndex);
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

                        WriteCellValue(writer, styledValues[styledRowOffset + c], cellValueKinds[c], dateTimeOffsetWriteStrategy, sharedStrings);
                    }

                    writer.Write("</row>");
                    rowIndex++;
                    styledRowOffset += columnCount;
                }
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
                if (sharedStrings == null) {
                    WritePlainInlineStringBufferedRows(writer, bufferedRows, rowCount, columnCount, startRowIndex, cellReferencePrefixes, ct);
                    return;
                }

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

            private static void WritePlainInlineStringBufferedRows(
                TextWriter writer,
                DirectBufferedRows bufferedRows,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
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
                            WriteStringCell(writer, text, validateLength: true);
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

            private static string GetWorksheetDimension(DirectDataSetSheetModel sheet) {
                var overlayCells = sheet.Metadata?.OverlayCells;
                if (overlayCells == null || overlayCells.Count == 0) {
                    return sheet.Range.Length == 0 ? "A1" : sheet.Range;
                }

                int lastRow = sheet.Table.RowCount + (sheet.IncludeHeaders ? 1 : 0);
                int lastColumn = sheet.Table.ColumnCount;
                for (int i = 0; i < overlayCells.Count; i++) {
                    lastRow = Math.Max(lastRow, overlayCells[i].Row);
                    lastColumn = Math.Max(lastColumn, overlayCells[i].Column);
                }

                return "A1:" + A1.CellReference(lastRow, lastColumn);
            }

            private static void WriteOverlayCells(
                TextWriter writer,
                IReadOnlyList<DirectOverlayCell>? overlayCells,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                int skipRowsAtOrBefore = 0) {
                if (overlayCells == null || overlayCells.Count == 0) {
                    return;
                }

                int row = -1;
                for (int i = 0; i < overlayCells.Count; i++) {
                    var cell = overlayCells[i];
                    if (cell.Row <= skipRowsAtOrBefore) {
                        continue;
                    }

                    if (cell.Row != row) {
                        if (row != -1) {
                            writer.Write("</row>");
                        }

                        row = cell.Row;
                        writer.Write("<row r=\"");
                        writer.Write(InvariantNumberText.Get(row));
                        writer.Write("\">");
                    }

                    WriteCell(
                        writer,
                        InvariantNumberText.Get(cell.Row),
                        "<c r=\"" + A1.ColumnIndexToLetters(cell.Column),
                        cell.Value,
                        CreateOverlayStyleAttribute(cell, stylePlan),
                        dateTimeOffsetWriteStrategy,
                        sharedStrings);
                }

                if (row != -1) {
                    writer.Write("</row>");
                }
            }

            private static Dictionary<int, IReadOnlyList<DirectOverlayCell>>? CreateOverlayCellsByRow(IReadOnlyList<DirectOverlayCell>? overlayCells) {
                if (overlayCells == null || overlayCells.Count == 0) {
                    return null;
                }

                return overlayCells
                    .GroupBy(static cell => cell.Row)
                    .ToDictionary(
                        static group => group.Key,
                        static group => (IReadOnlyList<DirectOverlayCell>)group.OrderBy(static cell => cell.Column).ToArray());
            }

            private static bool HasOverlayCellsAtOrBefore(IReadOnlyList<DirectOverlayCell>? overlayCells, int row) {
                if (overlayCells == null || overlayCells.Count == 0) {
                    return false;
                }

                for (int i = 0; i < overlayCells.Count; i++) {
                    if (overlayCells[i].Row <= row) {
                        return true;
                    }
                }

                return false;
            }

            private static bool WriteOverlayCellsForRow(
                TextWriter writer,
                IReadOnlyDictionary<int, IReadOnlyList<DirectOverlayCell>>? overlayCellsByRow,
                int row,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings) {
                if (overlayCellsByRow == null || !overlayCellsByRow.TryGetValue(row, out var rowOverlayCells)) {
                    return false;
                }

                string rowReference = InvariantNumberText.Get(row);
                for (int i = 0; i < rowOverlayCells.Count; i++) {
                    var cell = rowOverlayCells[i];
                    WriteCell(
                        writer,
                        rowReference,
                        "<c r=\"" + A1.ColumnIndexToLetters(cell.Column),
                        cell.Value,
                        CreateOverlayStyleAttribute(cell, stylePlan),
                        dateTimeOffsetWriteStrategy,
                        sharedStrings);
                }

                return true;
            }

            private static string? CreateOverlayStyleAttribute(DirectOverlayCell cell, DirectStylePlan stylePlan) {
                if (!string.IsNullOrWhiteSpace(cell.NumberFormat)) {
                    return stylePlan.GetStyleAttribute(cell.NumberFormat!);
                }

                return cell.StyleIndex == 0U ? " s=\"0\"" : null;
            }

            private static string[] CreateCellReferencePrefixes(int columnCount) {
                if ((uint)columnCount <= CachedCellReferencePrefixColumnLimit) {
                    return CellReferencePrefixCache[columnCount];
                }

                var columns = new string[columnCount];
                for (int i = 0; i < columnCount; i++) {
                    columns[i] = "<c r=\"" + A1.ColumnIndexToLetters(i + 1);
                }

                return columns;
            }

            private static string[][] CreateCellReferencePrefixCache() {
                var cache = new string[CachedCellReferencePrefixColumnLimit + 1][];
                cache[0] = Array.Empty<string>();
                for (int columnCount = 1; columnCount < cache.Length; columnCount++) {
                    var columns = new string[columnCount];
                    for (int i = 0; i < columnCount; i++) {
                        columns[i] = "<c r=\"" + A1.ColumnIndexToLetters(i + 1);
                    }

                    cache[columnCount] = columns;
                }

                return cache;
            }

            private static DirectColumnWritePlan CreateColumnWritePlan(
                DirectDataSetTableModel table,
                bool useCellValueNumberFormats,
                IReadOnlyList<string?>? columnNumberFormats,
                DirectStylePlan stylePlan) {
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

                    string? styleAttribute = null;
                    if (columnNumberFormats != null
                        && i < columnNumberFormats.Count
                        && columnNumberFormats[i] is string numberFormat
                        && !string.IsNullOrWhiteSpace(numberFormat)) {
                        styleAttribute = stylePlan.GetStyleAttribute(numberFormat);
                    }

                    styleAttribute ??= GetStyleAttribute(cellValueKind, useCellValueNumberFormats);
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
                if (dataType == typeof(DirectFormulaCellValue)) return DirectCellValueKind.Formula;
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
                int columnCount = sheet.Table.ColumnCount;
                var builder = new StringBuilder(512 + (columnCount * 72));
                string sheetIndexText = InvariantNumberText.Get(sheet.Index);
                builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                builder.Append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
                builder.Append(sheetIndexText);
                builder.Append("\" name=\"");
                AppendEscaped(builder, sheet.TableName!);
                builder.Append("\" displayName=\"");
                AppendEscaped(builder, sheet.TableName!);
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
                builder.Append(InvariantNumberText.Get(columnCount));
                builder.Append("\">");
                for (int i = 0; i < columnCount; i++) {
                    string columnIndexText = InvariantNumberText.Get(i + 1);
                    builder.Append("<tableColumn id=\"");
                    builder.Append(columnIndexText);
                    builder.Append("\" name=\"");
                    string columnName = sheet.Table.GetColumnName(i);
                    if (string.IsNullOrWhiteSpace(columnName)) {
                        builder.Append("Column");
                        builder.Append(columnIndexText);
                    } else {
                        AppendEscaped(builder, columnName);
                    }

                    builder.Append("\"/>");
                }

                builder.Append("</tableColumns><tableStyleInfo name=\"");
                builder.Append(sheet.TableStyle.ToString());
                builder.Append("\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>");
                WriteTextEntry(archive, "xl/tables/table" + sheetIndexText + ".xml", builder.ToString());
            }

            internal enum DirectCellValueKind {
                Object,
                Formula,
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

            internal sealed class DirectStylePlan {
                private const string DateTimeFormatCode = "yyyy-mm-dd hh:mm";
                private const string TimeFormatCode = "[h]:mm:ss";
                private readonly Dictionary<string, string> _styleAttributeByFormat;

                private DirectStylePlan(List<string> customNumberFormats, Dictionary<string, string> styleAttributeByFormat) {
                    CustomNumberFormats = customNumberFormats;
                    _styleAttributeByFormat = styleAttributeByFormat;
                }

                internal IReadOnlyList<string> CustomNumberFormats { get; }

                internal static DirectStylePlan Create(DirectDataSetWorkbookModel model) {
                    var customNumberFormats = new List<string>();
                    var styleAttributeByFormat = new Dictionary<string, string>(StringComparer.Ordinal);
                    styleAttributeByFormat[DateTimeFormatCode] = DateStyleAttribute;
                    styleAttributeByFormat[TimeFormatCode] = TimeStyleAttribute;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var formats = model.Sheets[sheetIndex].ColumnNumberFormats;
                        if (formats != null) {
                            for (int i = 0; i < formats.Count; i++) {
                                AddCustomNumberFormat(formats[i], customNumberFormats, styleAttributeByFormat);
                            }
                        }

                        var overlayCells = model.Sheets[sheetIndex].Metadata?.OverlayCells;
                        if (overlayCells != null) {
                            for (int i = 0; i < overlayCells.Count; i++) {
                                AddCustomNumberFormat(overlayCells[i].NumberFormat, customNumberFormats, styleAttributeByFormat);
                            }
                        }
                    }

                    return new DirectStylePlan(customNumberFormats, styleAttributeByFormat);
                }

                private static void AddCustomNumberFormat(string? format, List<string> customNumberFormats, Dictionary<string, string> styleAttributeByFormat) {
                    if (string.IsNullOrWhiteSpace(format) || styleAttributeByFormat.ContainsKey(format!)) {
                        return;
                    }

                    string styleAttribute = " s=\"" + InvariantNumberText.Get(5 + customNumberFormats.Count) + "\"";
                    styleAttributeByFormat.Add(format!, styleAttribute);
                    customNumberFormats.Add(format!);
                }

                internal string? GetStyleAttribute(string numberFormat)
                    => _styleAttributeByFormat.TryGetValue(numberFormat, out string? styleAttribute)
                        ? styleAttribute
                        : null;
            }

            internal sealed class ExtendedDirectWritePlan {
                internal ExtendedDirectWritePlan(
                    DirectDataSetWorkbookModel model,
                    DirectStylePlan stylePlan,
                    DirectColumnWritePlan[] columnWritePlans,
                    DirectSharedStringTable? sharedStrings) {
                    Model = model;
                    StylePlan = stylePlan;
                    ColumnWritePlans = columnWritePlans;
                    SharedStrings = sharedStrings;
                }

                internal DirectDataSetWorkbookModel Model { get; }

                internal DirectStylePlan StylePlan { get; }

                internal DirectColumnWritePlan[] ColumnWritePlans { get; }

                internal DirectSharedStringTable? SharedStrings { get; }

                internal bool HasSharedStrings => SharedStrings != null;
            }

            internal readonly struct DirectColumnWritePlan {
                internal DirectColumnWritePlan(DirectCellValueKind[] cellValueKinds, string?[]? styleAttributes, bool[]? valueStyleColumns) {
                    CellValueKinds = cellValueKinds;
                    StyleAttributes = styleAttributes;
                    ValueStyleColumns = valueStyleColumns;
                }

                internal DirectCellValueKind[] CellValueKinds { get; }

                internal string?[]? StyleAttributes { get; }

                internal bool[]? ValueStyleColumns { get; }
            }

            internal sealed class DirectSharedStringTable {
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

                    var stringCounts = new Dictionary<string, int>(StringComparer.Ordinal);
                    int totalStringReferences = 0;
                    int duplicateReferences = 0;
                    int sharedValueCount = 0;
                    int seenOnceCount = 0;
                    long totalStringCharacters = 0L;
                    long duplicateCharacters = 0L;
                    bool canCancel = ct.CanBeCanceled;
                    for (int sheetIndex = 0; sheetIndex < model.Sheets.Count; sheetIndex++) {
                        var sheet = model.Sheets[sheetIndex];
                        // Direct cell-value and dictionary-backed exports already use compact internal buffers
                        // or pay keyed lookup cost while writing; for these workloads the shared-string prepass
                        // tends to add CPU/allocation without enough package-size benefit.
                        if (sheet.Table.TryGetCellValueRows(out _)
                            || sheet.Table.TryGetExactDictionaryRows(out _)
                            || sheet.Table.TryGetDictionaryRows(out _)
                            || sheet.Table.TryGetLegacyDictionaryRows(out _)) {
                            return null;
                        }

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
                        } else if (sheet.Table.TryGetLegacyDictionaryRows(out _)) {
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

                    if (totalStringReferences < MinimumStringReferences || sharedValueCount == 0) {
                        return null;
                    }

                    if (duplicateReferences < MinimumDuplicateReferences && duplicateCharacters < MinimumDuplicateCharacters) {
                        return null;
                    }

                    var indexes = new Dictionary<string, int>(sharedValueCount, StringComparer.Ordinal);
                    var values = new string[sharedValueCount];
                    int nextIndex = 0;
                    int sharedStringReferences = 0;
                    foreach (var entry in stringCounts) {
                        int referenceCount = entry.Value < 0 ? -entry.Value : entry.Value;
                        if (referenceCount == 1 && entry.Value > 0) {
                            continue;
                        }

                        CoerceValueHelper.ValidateSharedStringLength(entry.Key, "value");
                        values[nextIndex] = entry.Key;
                        indexes[entry.Key] = nextIndex;
                        sharedStringReferences += referenceCount;
                        nextIndex++;
                    }

                    return new DirectSharedStringTable(indexes, values, sharedStringReferences);

                    void NoteString(string text, bool forceShared = false) {
                        totalStringReferences++;
                        totalStringCharacters += text.Length;
                        if (stringCounts.TryGetValue(text, out int count)) {
                            if (count == 1) {
                                seenOnceCount--;
                                sharedValueCount++;
                                stringCounts[text] = 2;
                            } else if (count > 1) {
                                stringCounts[text] = count + 1;
                            } else {
                                stringCounts[text] = count - 1;
                            }

                            duplicateReferences++;
                            duplicateCharacters += text.Length;
                            return;
                        }

                        if (forceShared) {
                            stringCounts.Add(text, -1);
                            sharedValueCount++;
                            return;
                        }

                        if (seenOnceCount >= MaximumSeenOnceCandidates) {
                            return;
                        }

                        stringCounts.Add(text, 1);
                        seenOnceCount++;
                    }

                    bool ShouldAbandonSharedStrings() {
                        return totalStringReferences >= MinimumEarlyUniqueHeavyStringReferences
                            && seenOnceCount >= MaximumSeenOnceCandidates
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

            private static DirectColumnWritePlan[] CreateColumnWritePlans(DirectDataSetWorkbookModel model, DirectStylePlan stylePlan, CancellationToken ct) {
                var plans = new DirectColumnWritePlan[model.Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < plans.Length; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    DirectDataSetSheetModel sheet = model.Sheets[i];
                    plans[i] = CreateColumnWritePlan(sheet.Table, sheet.UseCellValueNumberFormats, sheet.ColumnNumberFormats, stylePlan);
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
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct,
                IReadOnlyDictionary<int, IReadOnlyList<DirectOverlayCell>>? overlayCellsByRow = null) {
                if (overlayCellsByRow != null) {
                    WriteDirectValueRowsWithOverlayCells(writer, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings, ct, overlayCellsByRow);
                    return;
                }

                if (sheet.Table.TryGetExactDictionaryRows(out var exactDictionaryRows)) {
                    WriteExactDictionaryValueRows(writer, exactDictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetDictionaryRows(out var dictionaryRows)) {
                    WriteDictionaryValueRows(writer, dictionaryRows, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, sheet.Table.CreateColumnNameArray(), styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

                if (sheet.Table.TryGetLegacyDictionaryRows(out _)) {
                    WriteLegacyDictionaryValueRows(writer, sheet, rowCount, columnCount, startRowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, valueStyleColumns, dateTimeOffsetWriteStrategy, sharedStrings, ct);
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

            private static void WriteDirectValueRowsWithOverlayCells(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
                int rowCount,
                int columnCount,
                int startRowIndex,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                DirectCellValueKind[] cellValueKinds,
                bool[]? valueStyleColumns,
                DirectStylePlan stylePlan,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct,
                IReadOnlyDictionary<int, IReadOnlyList<DirectOverlayCell>> overlayCellsByRow) {
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

                        if (overlayCellsByRow.ContainsKey(rowIndex)) {
                            if (!rowStarted) {
                                writer.Write("<row r=\"");
                                writer.Write(rowReference);
                                writer.Write("\">");
                            }

                            WriteOverlayCellsForRow(writer, overlayCellsByRow, rowIndex, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings);
                            rowStarted = true;
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

                    WriteOverlayCellsForRow(writer, overlayCellsByRow, rowIndex, stylePlan, dateTimeOffsetWriteStrategy, sharedStrings);
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
                bool canCancel = ct.CanBeCanceled;
                int rowIndex = startRowIndex;
                if (!sheet.OmitBlankCells && valueStyleColumns == null) {
                    WriteFixedKindLegacyDictionaryRows(writer, sheet, rowCount, columnCount, rowIndex, cellReferencePrefixes, styleAttributes, cellValueKinds, dateTimeOffsetWriteStrategy, sharedStrings, ct);
                    return;
                }

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

            private static void WriteFixedKindLegacyDictionaryRows(
                TextWriter writer,
                DirectDataSetSheetModel sheet,
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
                        writer.Write("<row r=\"");
                        writer.Write(rowReference);
                        writer.Write("\">");
                        for (int c = 0; c < columnCount; c++) {
                            writer.Write(cellReferencePrefixes[c]);
                            writer.Write(rowReference);
                            writer.Write('"');
                            object? value = sheet.Table.GetValue(sourceRowIndex, c);
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

                        object? value = sheet.Table.GetValue(sourceRowIndex, c);
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
                    case DirectCellValueKind.Formula:
                        if (value is DirectFormulaCellValue formulaValue) {
                            WriteFormulaCellValue(writer, formulaValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Object:
                        if (value is DirectTypedCellValue typedValue) {
                            WriteTypedCellValue(writer, typedValue);
                            return;
                        }

                        break;
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
                    case DirectFormulaCellValue formulaValue:
                        WriteFormulaCellValue(writer, formulaValue);
                        return;
                    case DirectTypedCellValue typedValue:
                        WriteTypedCellValue(writer, typedValue);
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

            private static void WriteFormulaCellValue(TextWriter writer, string formula) {
                writer.Write("><f>");
                WriteEscaped(writer, formula);
                writer.Write("</f></c>");
            }

            private static void WriteFormulaCellValue(TextWriter writer, DirectFormulaCellValue formula) {
                if (!string.IsNullOrEmpty(formula.FormulaXml)) {
                    writer.Write('>');
                    writer.Write(formula.FormulaXml);
                    writer.Write("</c>");
                    return;
                }

                WriteFormulaCellValue(writer, formula.Formula);
            }

            private static void WriteTypedCellValue(TextWriter writer, DirectTypedCellValue typed) {
                writer.Write(" t=\"");
                WriteEscaped(writer, typed.DataType);
                writer.Write("\">");
                if (!string.IsNullOrEmpty(typed.InlineStringXml)) {
                    writer.Write(typed.InlineStringXml);
                } else if (typed.Value != null) {
                    writer.Write("<v>");
                    WriteEscaped(writer, typed.Value);
                    writer.Write("</v>");
                } else {
                    writer.Write("<v/>");
                }

                writer.Write("</c>");
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
                if ((uint)sharedStringIndex < (uint)SharedStringCellCache.Length) {
                    string? cached = SharedStringCellCache[sharedStringIndex];
                    if (cached == null) {
                        cached = " t=\"s\"><v>" + InvariantNumberText.Get(sharedStringIndex) + "</v></c>";
                        SharedStringCellCache[sharedStringIndex] = cached;
                    }

                    writer.Write(cached);
                    return;
                }

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
                if (TryWriteCachedRawNonNegativeIntegerCell(writer, value)) {
                    return;
                }

                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, long value) {
                if (TryWriteCachedRawNonNegativeIntegerCell(writer, value)) {
                    return;
                }

                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static void WriteRawValueCell(TextWriter writer, ulong value) {
                if (TryWriteCachedRawNonNegativeIntegerCell(writer, value)) {
                    return;
                }

                writer.Write(" t=\"n\"><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static bool TryWriteCachedRawNonNegativeIntegerCell(TextWriter writer, int value) {
                var cache = RawNonNegativeIntegerCellCache;
                if ((uint)value >= (uint)cache.Length) {
                    return false;
                }

                writer.Write(cache[value]);
                return true;
            }

            private static bool TryWriteCachedRawNonNegativeIntegerCell(TextWriter writer, long value) {
                var cache = RawNonNegativeIntegerCellCache;
                if ((ulong)value >= (ulong)cache.Length) {
                    return false;
                }

                writer.Write(cache[(int)value]);
                return true;
            }

            private static bool TryWriteCachedRawNonNegativeIntegerCell(TextWriter writer, ulong value) {
                var cache = RawNonNegativeIntegerCellCache;
                if (value >= (ulong)cache.Length) {
                    return false;
                }

                writer.Write(cache[(int)value]);
                return true;
            }

            private static string[] CreateRawNonNegativeIntegerCellCache() {
                var cache = new string[CachedRawIntegerCellLimit];
                for (int i = 0; i < cache.Length; i++) {
                    cache[i] = " t=\"n\"><v>" + InvariantNumberText.Get(i) + "</v></c>";
                }

                return cache;
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
#if NET6_0_OR_GREATER
                int byteCount = Utf8NoBom.GetByteCount(text);
                if (byteCount <= StackallocTextEntryByteLimit) {
                    Span<byte> stackBytes = stackalloc byte[byteCount];
                    int written = Utf8NoBom.GetBytes(text.AsSpan(), stackBytes);
                    stream.Write(stackBytes.Slice(0, written));
                    return;
                }
#endif
                byte[] bytes = Utf8NoBom.GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
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
                    if (!IsInvalidXmlControl(current) && !IsXmlTextEscape(current)) {
                        continue;
                    }

                    if (i > start) {
                        WriteSlice(writer, value, start, i - start);
                    }

                    if (!IsInvalidXmlControl(current)) {
                        WriteEscapedTextCharacter(writer, current);
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

            private static bool IsXmlTextEscape(char value)
                => value is '&' or '<' or '>';

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

            private static void WriteEscapedTextCharacter(TextWriter writer, char value) {
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
                }
            }

        }

    }
}
