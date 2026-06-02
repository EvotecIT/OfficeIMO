using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private const int XmlWriterBufferSize = 65536;
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

            internal static bool TryCreateExtendedWritePlan(DirectDataSetWorkbookModel model, CancellationToken ct, out ExtendedDirectWritePlan? plan, bool disableSharedStrings = false) {
                DirectStylePlan stylePlan = DirectStylePlan.Create(model);
                DirectColumnWritePlan[] columnWritePlans = CreateColumnWritePlans(model, stylePlan, ct);
                var sharedStrings = disableSharedStrings ? null : DirectSharedStringTable.Create(model, columnWritePlans, ct);

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

        }

    }
}
