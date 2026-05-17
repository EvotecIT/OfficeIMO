using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO.Compression;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool TryWriteSimpleWorkbookPackage(Stream destination) {
            if (destination == null || !destination.CanWrite || !destination.CanSeek) {
                return false;
            }

            if (_packagePropertiesDirty || _unchangedPackageBytes != null || HasCalculationSaveWork()) {
                return false;
            }

            if (!FastWorkbookPackageModel.TryCreate(_spreadSheetDocument, out var model)) {
                return false;
            }

            PrepareDestinationStreamForWrite(destination);
            FastWorkbookPackageWriter.Write(destination, model);

            destination.Flush();
            destination.Seek(0, SeekOrigin.Begin);
            _packageDirty = false;
            _packagePropertiesDirty = false;
            _requiresSavePreflight = false;
            _unchangedPackageBytes = null;
            return true;
        }

        private static class FastWorkbookPackageWriter {
            internal static void Write(Stream destination, FastWorkbookPackageModel model) {
                using (var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true)) {
                    WriteContentTypesEntry(archive, model.HasStyles, model.Worksheets.Count, model.Tables.Count);
                    WriteTextEntry(archive, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                        "</Relationships>");
                    WriteWorkbookEntry(archive, model.Worksheets);
                    WriteWorkbookRelationshipsEntry(archive, model.Worksheets, model.HasStyles);
                    if (model.Stylesheet != null) {
                        WriteTextEntry(archive, "xl/styles.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + model.Stylesheet.OuterXml);
                    }

                    foreach (var worksheet in model.Worksheets) {
                        WriteWorksheetEntry(archive, worksheet, model.SharedStrings);
                        if (worksheet.TablePartPaths.Count > 0) {
                            WriteWorksheetRelationshipsEntry(archive, worksheet.RelationshipsPath, worksheet.TablePartPaths);
                        }
                    }

                    for (int i = 0; i < model.Tables.Count; i++) {
                        WriteTextEntry(
                            archive,
                            string.Format(CultureInfo.InvariantCulture, "xl/tables/table{0}.xml", i + 1),
                            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + model.Tables[i].OuterXml);
                    }
                }
            }
        }

        private sealed class FastWorkbookPackageModel {
            private FastWorkbookPackageModel(
                IReadOnlyList<FastWorksheetPackageModel> worksheets,
                Stylesheet? stylesheet,
                IReadOnlyList<string>? sharedStrings,
                IReadOnlyList<Table> tables) {
                Worksheets = worksheets;
                Stylesheet = stylesheet;
                SharedStrings = sharedStrings;
                Tables = tables;
            }

            internal IReadOnlyList<FastWorksheetPackageModel> Worksheets { get; }

            internal Stylesheet? Stylesheet { get; }

            internal bool HasStyles => Stylesheet != null;

            internal IReadOnlyList<string>? SharedStrings { get; }

            internal IReadOnlyList<Table> Tables { get; }

            internal static bool TryCreate(SpreadsheetDocument document, out FastWorkbookPackageModel model) {
                model = null!;

                var workbookPart = document.WorkbookPart;
                if (workbookPart?.Workbook?.Sheets == null) {
                    return false;
                }

                var sheets = workbookPart.Workbook.Sheets.OfType<Sheet>().ToList();
                if (sheets.Count == 0 || sheets.Any(sheet => sheet.Id == null)) {
                    return false;
                }

                if (workbookPart.Workbook.Descendants<DefinedName>().Any()
                    || workbookPart.CalculationChainPart != null) {
                    return false;
                }

                if (workbookPart.GetPartsOfType<ThemePart>().Any()) {
                    return false;
                }

                var worksheets = new List<FastWorksheetPackageModel>(sheets.Count);
                var tables = new List<Table>();
                int tableIndex = 1;
                for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
                    var sheet = sheets[sheetIndex];
                    if (workbookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart) {
                        return false;
                    }

                    var worksheet = worksheetPart.Worksheet;
                    if (worksheet == null || !CanWriteWorksheet(worksheetPart, worksheet)) {
                        return false;
                    }

                    var tablePartPaths = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var tableDefinition in worksheetPart.TableDefinitionParts) {
                        var table = tableDefinition.Table;
                        if (table == null) {
                            return false;
                        }

                        tables.Add(table);
                        string relId = worksheetPart.GetIdOfPart(tableDefinition);
                        tablePartPaths[relId] = string.Format(CultureInfo.InvariantCulture, "../tables/table{0}.xml", tableIndex);
                        tableIndex++;
                    }

                    worksheets.Add(new FastWorksheetPackageModel(
                        sheet.Name?.Value ?? "Sheet" + (sheetIndex + 1).ToString(CultureInfo.InvariantCulture),
                        sheet.SheetId?.Value ?? (uint)(sheetIndex + 1),
                        "rId" + (sheetIndex + 1).ToString(CultureInfo.InvariantCulture),
                        string.Format(CultureInfo.InvariantCulture, "xl/worksheets/sheet{0}.xml", sheetIndex + 1),
                        string.Format(CultureInfo.InvariantCulture, "xl/worksheets/_rels/sheet{0}.xml.rels", sheetIndex + 1),
                        worksheet,
                        tablePartPaths));
                }

                var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                    .Elements<SharedStringItem>()
                    .Select(item => item.InnerText ?? string.Empty)
                    .ToList();

                model = new FastWorkbookPackageModel(
                    worksheets,
                    workbookPart.WorkbookStylesPart?.Stylesheet,
                    sharedStrings,
                    tables);
                return true;
            }

            private static bool CanWriteWorksheet(WorksheetPart worksheetPart, Worksheet worksheet) {
                return CanWriteSimpleWorksheet(worksheetPart, worksheet);
            }
        }

        private sealed class FastWorksheetPackageModel {
            internal FastWorksheetPackageModel(
                string sheetName,
                uint sheetId,
                string workbookRelationshipId,
                string worksheetPath,
                string relationshipsPath,
                Worksheet worksheet,
                IReadOnlyDictionary<string, string> tablePartPaths) {
                SheetName = sheetName;
                SheetId = sheetId;
                WorkbookRelationshipId = workbookRelationshipId;
                WorksheetPath = worksheetPath;
                RelationshipsPath = relationshipsPath;
                Worksheet = worksheet;
                TablePartPaths = tablePartPaths;
            }

            internal string SheetName { get; }

            internal uint SheetId { get; }

            internal string WorkbookRelationshipId { get; }

            internal string WorksheetPath { get; }

            internal string RelationshipsPath { get; }

            internal Worksheet Worksheet { get; }

            internal IReadOnlyDictionary<string, string> TablePartPaths { get; }
        }

        private static bool CanWriteSimpleWorksheet(WorksheetPart worksheetPart, Worksheet worksheet) {
            if (worksheetPart.DrawingsPart != null
                || worksheetPart.WorksheetCommentsPart != null
                || worksheetPart.HyperlinkRelationships.Any()
                || worksheetPart.ExternalRelationships.Any()) {
                return false;
            }

            foreach (var child in worksheet.ChildElements) {
                if (child is not SheetProperties
                    && child is not SheetDimension
                    && child is not SheetViews
                    && child is not SheetFormatProperties
                    && child is not Columns
                    && child is not SheetData
                    && child is not SheetProtection
                    && child is not AutoFilter
                    && child is not SortState
                    && child is not MergeCells
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                    && child is not DataValidations
                    && child is not TableParts) {
                    return false;
                }

                if (child.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    return false;
                }
            }

            var tableParts = worksheet.Elements<TableParts>().ToList();
            if (tableParts.Count > 1) {
                return false;
            }

            var tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
            var relationshipIds = new HashSet<string>(tableDefinitionParts.Select(worksheetPart.GetIdOfPart), StringComparer.Ordinal);
            var worksheetTablePartIds = tableParts.Count == 0
                ? new List<string>()
                : tableParts[0].Elements<TablePart>()
                    .Select(part => part.Id?.Value)
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Select(id => id!)
                    .ToList();

            if (worksheetTablePartIds.Count != tableDefinitionParts.Count
                || worksheetTablePartIds.Any(id => !relationshipIds.Contains(id))) {
                return false;
            }

            foreach (var tableDefinitionPart in tableDefinitionParts) {
                var table = tableDefinitionPart.Table;
                if (table == null
                    || table.Reference == null
                    || table.TableColumns == null
                    || table.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    return false;
                }
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return true;
            }

            foreach (var row in sheetData.Elements<Row>()) {
                if (!IsSimpleRow(row)) {
                    return false;
                }

                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellFormula != null || cell.InlineString != null) {
                        return false;
                    }

                    var dataType = cell.DataType?.Value;
                    if (dataType != null
                        && dataType != CellValues.Number
                        && dataType != CellValues.SharedString
                        && dataType != CellValues.String
                        && dataType != CellValues.Boolean) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsSimpleRow(Row row) {
            foreach (var attribute in row.GetAttributes()) {
                if (!string.Equals(attribute.LocalName, "r", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return row.CustomFormat?.Value != true && row.StyleIndex == null;
        }

        private static void WriteContentTypesEntry(ZipArchive archive, bool hasStyles, int worksheetCount, int tableCount) {
            var builder = new System.Text.StringBuilder(512 + worksheetCount * 160 + tableCount * 160);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 1; i <= worksheetCount; i++) {
                builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                builder.Append(i.ToString(CultureInfo.InvariantCulture));
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }

            if (hasStyles) {
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            }

            for (int i = 1; i <= tableCount; i++) {
                builder.Append("<Override PartName=\"/xl/tables/table");
                builder.Append(i.ToString(CultureInfo.InvariantCulture));
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            }

            builder.Append("</Types>");
            WriteTextEntry(archive, "[Content_Types].xml", builder.ToString());
        }

        private static void WriteWorkbookEntry(ZipArchive archive, IReadOnlyList<FastWorksheetPackageModel> worksheets) {
            var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateFastXmlWriter(stream);
            writer.WriteStartDocument();
            writer.WriteStartElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            writer.WriteStartElement("sheets");
            foreach (var worksheet in worksheets) {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", worksheet.SheetName);
                writer.WriteAttributeString("sheetId", worksheet.SheetId.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", worksheet.WorkbookRelationshipId);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteWorkbookRelationshipsEntry(ZipArchive archive, IReadOnlyList<FastWorksheetPackageModel> worksheets, bool hasStyles) {
            var builder = new System.Text.StringBuilder(384 + worksheets.Count * 180);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var worksheet in worksheets) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, worksheet.WorkbookRelationshipId);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"");
                AppendXmlEscaped(builder, worksheet.WorksheetPath.Substring("xl/".Length));
                builder.Append("\"/>");
            }

            if (hasStyles) {
                builder.Append("<Relationship Id=\"rId");
                builder.Append((worksheets.Count + 1).ToString(CultureInfo.InvariantCulture));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", builder.ToString());
        }

        private static void WriteWorksheetRelationshipsEntry(ZipArchive archive, string relationshipsPath, IReadOnlyDictionary<string, string> tablePartPaths) {
            var builder = new System.Text.StringBuilder(160 + tablePartPaths.Count * 180);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var item in tablePartPaths.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, item.Key);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"");
                AppendXmlEscaped(builder, item.Value);
                builder.Append("\"/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, relationshipsPath, builder.ToString());
        }

        private static void WriteWorksheetEntry(ZipArchive archive, FastWorksheetPackageModel model, IReadOnlyList<string>? sharedStrings) {
            var entry = archive.CreateEntry(model.WorksheetPath, CompressionLevel.Fastest);
            var worksheet = model.Worksheet;
            string dimension = worksheet.SheetDimension?.Reference?.Value ?? ExcelSheet.ComputeSheetDimensionReference(worksheet);
            var builder = new System.Text.StringBuilder(4096);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            if (model.TablePartPaths.Count > 0) {
                builder.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
            }

            builder.Append(">");
            AppendOptionalElement(builder, worksheet.GetFirstChild<SheetProperties>());

            builder.Append("<dimension ref=\"");
            AppendXmlEscaped(builder, dimension);
            builder.Append("\"/>");

            AppendOptionalElement(builder, worksheet.GetFirstChild<SheetViews>());
            AppendOptionalElement(builder, worksheet.GetFirstChild<SheetFormatProperties>());

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                AppendColumns(builder, columns);
            }

            builder.Append("<sheetData>");

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                foreach (var row in sheetData.Elements<Row>()) {
                    builder.Append("<row");
                    if (row.RowIndex != null) {
                        builder.Append(" r=\"");
                        builder.Append(row.RowIndex.Value.ToString(CultureInfo.InvariantCulture));
                        builder.Append('"');
                    }

                    builder.Append('>');

                    foreach (var cell in row.Elements<Cell>()) {
                        AppendSimpleCell(builder, cell, sharedStrings);
                    }

                    builder.Append("</row>");
                }
            }

            builder.Append("</sheetData>");
            AppendOptionalElement(builder, worksheet.GetFirstChild<SheetProtection>());

            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                builder.Append(autoFilter.OuterXml);
            }

            AppendOptionalElement(builder, worksheet.GetFirstChild<SortState>());
            AppendOptionalElement(builder, worksheet.GetFirstChild<MergeCells>());
            AppendOptionalElements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>(builder, worksheet);
            AppendOptionalElement(builder, worksheet.GetFirstChild<DataValidations>());

            var tableParts = worksheet.GetFirstChild<TableParts>();
            if (tableParts != null && model.TablePartPaths.Count > 0) {
                builder.Append("<tableParts count=\"");
                builder.Append(model.TablePartPaths.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append("\">");
                foreach (var tablePart in tableParts.Elements<TablePart>()) {
                    string? id = tablePart.Id?.Value;
                    if (id == null || !model.TablePartPaths.ContainsKey(id)) {
                        continue;
                    }

                    builder.Append("<tablePart r:id=\"");
                    AppendXmlEscaped(builder, id);
                    builder.Append("\"/>");
                }

                builder.Append("</tableParts>");
            }

            builder.Append("</worksheet>");
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            writer.Write(builder.ToString());
        }

        private static void AppendOptionalElement(System.Text.StringBuilder builder, OpenXmlElement? element) {
            if (element != null) {
                builder.Append(element.OuterXml);
            }
        }

        private static void AppendOptionalElements<TElement>(System.Text.StringBuilder builder, OpenXmlElement parent)
            where TElement : OpenXmlElement {
            foreach (var element in parent.Elements<TElement>()) {
                builder.Append(element.OuterXml);
            }
        }

        private static void AppendColumns(System.Text.StringBuilder builder, Columns columns) {
            builder.Append("<cols>");
            foreach (var column in columns.Elements<Column>()) {
                builder.Append("<col");
                AppendUIntAttribute(builder, "min", column.Min);
                AppendUIntAttribute(builder, "max", column.Max);
                if (column.Width != null) {
                    builder.Append(" width=\"");
                    builder.Append(column.Width.Value.ToString(CultureInfo.InvariantCulture));
                    builder.Append('"');
                }

                AppendBooleanAttribute(builder, "bestFit", column.BestFit);
                AppendBooleanAttribute(builder, "customWidth", column.CustomWidth);
                AppendBooleanAttribute(builder, "hidden", column.Hidden);
                AppendUIntAttribute(builder, "style", column.Style);
                AppendByteAttribute(builder, "outlineLevel", column.OutlineLevel);
                AppendBooleanAttribute(builder, "collapsed", column.Collapsed);
                builder.Append("/>");
            }

            builder.Append("</cols>");
        }

        private static void AppendSimpleCell(System.Text.StringBuilder builder, Cell cell, IReadOnlyList<string>? sharedStrings) {
            string? text = cell.CellValue?.Text;
            var dataType = cell.DataType?.Value;
            if (dataType == CellValues.SharedString) {
                text = TryResolveSharedString(text, sharedStrings);
                dataType = CellValues.String;
            }

            builder.Append("<c");
            if (cell.CellReference != null) {
                builder.Append(" r=\"");
                AppendXmlEscaped(builder, cell.CellReference.Value ?? string.Empty);
                builder.Append('"');
            }

            if (cell.StyleIndex != null) {
                builder.Append(" s=\"");
                builder.Append(cell.StyleIndex.Value.ToString(CultureInfo.InvariantCulture));
                builder.Append('"');
            }

            if (dataType == CellValues.String) {
                builder.Append(" t=\"str\"");
            } else if (dataType == CellValues.Boolean) {
                builder.Append(" t=\"b\"");
            }

            builder.Append("><v>");
            AppendXmlEscaped(builder, text ?? string.Empty);
            builder.Append("</v></c>");
        }

        private static string TryResolveSharedString(string? rawIndex, IReadOnlyList<string>? sharedStrings) {
            if (sharedStrings != null
                && int.TryParse(rawIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                && index >= 0
                && index < sharedStrings.Count) {
                return sharedStrings[index];
            }

            return rawIndex ?? string.Empty;
        }

        private static void WriteTextEntry(ZipArchive archive, string path, string text) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            writer.Write(text);
        }

        private static void AppendUIntAttribute(System.Text.StringBuilder builder, string name, UInt32Value? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            builder.Append(value.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        private static void AppendByteAttribute(System.Text.StringBuilder builder, string name, ByteValue? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            builder.Append(value.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        private static void AppendBooleanAttribute(System.Text.StringBuilder builder, string name, BooleanValue? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            builder.Append(value.Value ? '1' : '0');
            builder.Append('"');
        }

        private static XmlWriter CreateFastXmlWriter(Stream stream) =>
            XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                CloseOutput = false,
                Indent = false,
                OmitXmlDeclaration = false
            });

        private static void AppendXmlEscaped(System.Text.StringBuilder builder, string text) {
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
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
