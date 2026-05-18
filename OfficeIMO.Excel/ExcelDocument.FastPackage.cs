using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO.Compression;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool TryWriteSimpleWorkbookPackage(Stream destination, ExcelSaveOptions? options, bool updateDocumentState, out string? skipReason) {
            skipReason = null;

            if (destination == null || !destination.CanWrite || !destination.CanSeek) {
                skipReason = "Destination stream must be writable and seekable.";
                return false;
            }

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (_unchangedPackageBytes != null) {
                skipReason = "An unchanged package payload is already available.";
                return false;
            }

            if (_packageContentTypesKnownNormalized && !_simplePackageContentKnown) {
                skipReason = "Workbook was loaded or previously finalized; standard save preserves package metadata and relationships.";
                return false;
            }

            if (HasCalculationSaveWork()) {
                skipReason = "Calculation save work is pending.";
                return false;
            }

            if (!FastWorkbookPackageModel.TryCreate(_spreadSheetDocument, out var model, out string? modelSkipReason)) {
                skipReason = modelSkipReason ?? "Workbook contains parts or worksheet features outside the simple package writer surface.";
                return false;
            }

            PrepareDestinationStreamForWrite(destination);
            FastWorkbookPackageWriter.Write(destination, model);

            destination.Flush();
            destination.Seek(0, SeekOrigin.Begin);
            if (updateDocumentState) {
                _packageDirty = false;
                _packagePropertiesDirty = false;
                _requiresSavePreflight = false;
                _unchangedPackageBytes = null;
                _packageContentTypesKnownNormalized = true;
                _simplePackageContentKnown = true;
            }

            return true;
        }

        private static class FastWorkbookPackageWriter {
            internal static void Write(Stream destination, FastWorkbookPackageModel model) {
                using (var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true)) {
                    WriteContentTypesEntry(archive, model.HasStyles, model.HasSharedStrings, model.Worksheets.Count, model.Tables.Count);
                    WriteTextEntry(archive, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                        "</Relationships>");
                    WriteCorePropertiesEntry(archive);
                    WriteAppPropertiesEntry(archive);
                    WriteWorkbookEntry(archive, model);
                    WriteWorkbookRelationshipsEntry(archive, model.Worksheets, model.HasStyles, model.HasSharedStrings);
                    if (model.Stylesheet != null) {
                        WriteTextEntry(archive, "xl/styles.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + model.Stylesheet.OuterXml);
                    }

                    if (model.HasSharedStrings && model.SharedStrings != null) {
                        WriteSharedStringsEntry(archive, model.SharedStrings);
                    }

                    foreach (var worksheet in model.Worksheets) {
                        WriteWorksheetEntry(archive, worksheet);
                        if (worksheet.HasRelationships) {
                            WriteWorksheetRelationshipsEntry(archive, worksheet);
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
                SharedStringTable? sharedStrings,
                IReadOnlyList<Table> tables,
                FileVersion? fileVersion,
                FileSharing? fileSharing,
                WorkbookProperties? workbookProperties,
                WorkbookProtection? workbookProtection,
                BookViews? bookViews,
                DefinedNames? definedNames,
                CalculationProperties? calculationProperties) {
                Worksheets = worksheets;
                Stylesheet = stylesheet;
                SharedStrings = sharedStrings;
                Tables = tables;
                FileVersion = fileVersion;
                FileSharing = fileSharing;
                WorkbookProperties = workbookProperties;
                WorkbookProtection = workbookProtection;
                BookViews = bookViews;
                DefinedNames = definedNames;
                CalculationProperties = calculationProperties;
            }

            internal IReadOnlyList<FastWorksheetPackageModel> Worksheets { get; }

            internal Stylesheet? Stylesheet { get; }

            internal bool HasStyles => Stylesheet != null;

            internal SharedStringTable? SharedStrings { get; }

            internal bool HasSharedStrings => SharedStrings != null && SharedStrings.Elements<SharedStringItem>().Any();

            internal IReadOnlyList<Table> Tables { get; }

            internal FileVersion? FileVersion { get; }

            internal FileSharing? FileSharing { get; }

            internal WorkbookProperties? WorkbookProperties { get; }

            internal WorkbookProtection? WorkbookProtection { get; }

            internal BookViews? BookViews { get; }

            internal DefinedNames? DefinedNames { get; }

            internal CalculationProperties? CalculationProperties { get; }

            internal static bool TryCreate(SpreadsheetDocument document, out FastWorkbookPackageModel model, out string? skipReason) {
                model = null!;
                skipReason = null;

                var workbookPart = document.WorkbookPart;
                if (workbookPart?.Workbook?.Sheets == null) {
                    skipReason = "Workbook is missing sheets.";
                    return false;
                }

                var sheets = workbookPart.Workbook.Sheets.OfType<Sheet>().ToList();
                if (sheets.Count == 0 || sheets.Any(sheet => sheet.Id == null)) {
                    skipReason = "Workbook has no sheets or has sheets without relationships.";
                    return false;
                }

                if (workbookPart.CalculationChainPart != null) {
                    skipReason = "Workbook contains a calculation chain part.";
                    return false;
                }

                var unsupportedWorkbookChild = workbookPart.Workbook.ChildElements
                    .FirstOrDefault(child => child is not DocumentFormat.OpenXml.Spreadsheet.FileVersion
                        && child is not DocumentFormat.OpenXml.Spreadsheet.FileSharing
                        && child is not DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties
                        && child is not DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection
                        && child is not DocumentFormat.OpenXml.Spreadsheet.BookViews
                        && child is not DocumentFormat.OpenXml.Spreadsheet.Sheets
                        && child is not DocumentFormat.OpenXml.Spreadsheet.DefinedNames
                        && child is not DocumentFormat.OpenXml.Spreadsheet.CalculationProperties);
                if (unsupportedWorkbookChild != null) {
                    skipReason = "Workbook contains unsupported workbook-level element '" + unsupportedWorkbookChild.LocalName + "'.";
                    return false;
                }

                foreach (var child in workbookPart.Workbook.ChildElements) {
                    if (child.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Workbook contains unknown Open XML elements.";
                        return false;
                    }
                }

                var definedNames = workbookPart.Workbook.GetFirstChild<DefinedNames>();
                if (workbookPart.GetPartsOfType<ThemePart>().Any()) {
                    skipReason = "Workbook contains a theme part.";
                    return false;
                }

                var worksheets = new List<FastWorksheetPackageModel>(sheets.Count);
                var tables = new List<Table>();
                int tableIndex = 1;
                for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
                    var sheet = sheets[sheetIndex];
                    if (workbookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart) {
                        skipReason = "Workbook sheet relationship does not target a worksheet part.";
                        return false;
                    }

                    var worksheet = worksheetPart.Worksheet;
                    if (worksheet == null) {
                        skipReason = "Worksheet part is missing worksheet XML.";
                        return false;
                    }

                    if (!CanWriteWorksheet(worksheetPart, worksheet, out skipReason)) {
                        return false;
                    }

                    var tablePartPaths = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var tableDefinition in worksheetPart.TableDefinitionParts) {
                        var table = tableDefinition.Table;
                        if (table == null) {
                            skipReason = "Worksheet table definition is missing table XML.";
                            return false;
                        }

                        tables.Add(table);
                        string relId = worksheetPart.GetIdOfPart(tableDefinition);
                        tablePartPaths[relId] = string.Format(CultureInfo.InvariantCulture, "../tables/table{0}.xml", tableIndex);
                        tableIndex++;
                    }

                    var hyperlinkRelationships = worksheetPart.HyperlinkRelationships
                        .Select(relationship => new FastHyperlinkRelationshipModel(
                            relationship.Id,
                            relationship.Uri.ToString(),
                            relationship.IsExternal))
                        .ToList();

                    worksheets.Add(new FastWorksheetPackageModel(
                        sheet.Name?.Value ?? "Sheet" + (sheetIndex + 1).ToString(CultureInfo.InvariantCulture),
                        sheet.SheetId?.Value ?? (uint)(sheetIndex + 1),
                        GetSheetStateText(sheet),
                        "rId" + (sheetIndex + 1).ToString(CultureInfo.InvariantCulture),
                        string.Format(CultureInfo.InvariantCulture, "xl/worksheets/sheet{0}.xml", sheetIndex + 1),
                        string.Format(CultureInfo.InvariantCulture, "xl/worksheets/_rels/sheet{0}.xml.rels", sheetIndex + 1),
                        worksheet,
                        tablePartPaths,
                        hyperlinkRelationships));
                }

                var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStrings != null
                    && sharedStrings.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    skipReason = "Workbook shared strings contain unknown Open XML elements.";
                    return false;
                }

                model = new FastWorkbookPackageModel(
                    worksheets,
                    workbookPart.WorkbookStylesPart?.Stylesheet,
                    sharedStrings!,
                    tables,
                    workbookPart.Workbook.GetFirstChild<FileVersion>(),
                    workbookPart.Workbook.GetFirstChild<FileSharing>(),
                    workbookPart.Workbook.GetFirstChild<WorkbookProperties>(),
                    workbookPart.Workbook.GetFirstChild<WorkbookProtection>(),
                    workbookPart.Workbook.GetFirstChild<BookViews>(),
                    definedNames,
                    workbookPart.Workbook.GetFirstChild<CalculationProperties>());
                return true;
            }

            private static bool CanWriteWorksheet(WorksheetPart worksheetPart, Worksheet worksheet, out string? skipReason) {
                return CanWriteSimpleWorksheet(worksheetPart, worksheet, out skipReason);
            }

            private static string? GetSheetStateText(Sheet sheet) {
                if (sheet.State == null) {
                    return null;
                }

                if (sheet.State.Value == SheetStateValues.Hidden) {
                    return "hidden";
                }

                if (sheet.State.Value == SheetStateValues.VeryHidden) {
                    return "veryHidden";
                }

                if (sheet.State.Value == SheetStateValues.Visible) {
                    return "visible";
                }

                return sheet.State.InnerText;
            }
        }

        private sealed class FastWorksheetPackageModel {
            internal FastWorksheetPackageModel(
                string sheetName,
                uint sheetId,
                string? sheetState,
                string workbookRelationshipId,
                string worksheetPath,
                string relationshipsPath,
                Worksheet worksheet,
                IReadOnlyDictionary<string, string> tablePartPaths,
                IReadOnlyList<FastHyperlinkRelationshipModel> hyperlinkRelationships) {
                SheetName = sheetName;
                SheetId = sheetId;
                SheetState = sheetState;
                WorkbookRelationshipId = workbookRelationshipId;
                WorksheetPath = worksheetPath;
                RelationshipsPath = relationshipsPath;
                Worksheet = worksheet;
                TablePartPaths = tablePartPaths;
                HyperlinkRelationships = hyperlinkRelationships;
            }

            internal string SheetName { get; }

            internal uint SheetId { get; }

            internal string? SheetState { get; }

            internal string WorkbookRelationshipId { get; }

            internal string WorksheetPath { get; }

            internal string RelationshipsPath { get; }

            internal Worksheet Worksheet { get; }

            internal IReadOnlyDictionary<string, string> TablePartPaths { get; }

            internal IReadOnlyList<FastHyperlinkRelationshipModel> HyperlinkRelationships { get; }

            internal bool HasRelationships => TablePartPaths.Count > 0 || HyperlinkRelationships.Count > 0;
        }

        private sealed class FastHyperlinkRelationshipModel {
            internal FastHyperlinkRelationshipModel(string id, string target, bool isExternal) {
                Id = id;
                Target = target;
                IsExternal = isExternal;
            }

            internal string Id { get; }

            internal string Target { get; }

            internal bool IsExternal { get; }
        }

        private static bool CanWriteSimpleWorksheet(WorksheetPart worksheetPart, Worksheet worksheet, out string? skipReason) {
            skipReason = null;

            if (worksheetPart.DrawingsPart != null) {
                skipReason = "Worksheet contains drawings.";
                return false;
            }

            if (worksheetPart.WorksheetCommentsPart != null) {
                skipReason = "Worksheet contains comments.";
                return false;
            }

            if (worksheetPart.ExternalRelationships.Any()) {
                skipReason = "Worksheet contains external relationships.";
                return false;
            }

            foreach (var child in worksheet.ChildElements) {
                if (child is not SheetProperties
                    && child is not SheetDimension
                    && child is not SheetViews
                    && child is not SheetFormatProperties
                    && child is not Columns
                    && child is not SheetData
                    && child is not SheetCalculationProperties
                    && child is not SheetProtection
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges
                    && child is not Scenarios
                    && child is not AutoFilter
                    && child is not SortState
                    && child is not MergeCells
                    && child is not PhoneticProperties
                    && child is not Hyperlinks
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                    && child is not DataValidations
                    && child is not PrintOptions
                    && child is not PageMargins
                    && child is not PageSetup
                    && child is not HeaderFooter
                    && child is not RowBreaks
                    && child is not ColumnBreaks
                    && child is not CellWatches
                    && child is not DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors
                    && child is not TableParts) {
                    skipReason = "Worksheet contains unsupported element '" + child.LocalName + "'.";
                    return false;
                }

                if (child.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    skipReason = "Worksheet contains unknown Open XML elements.";
                    return false;
                }
            }

            var tableParts = worksheet.Elements<TableParts>().ToList();
            if (tableParts.Count > 1) {
                skipReason = "Worksheet contains multiple tableParts elements.";
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
                skipReason = "Worksheet table relationships do not match tableParts entries.";
                return false;
            }

            var hyperlinkRelationships = worksheetPart.HyperlinkRelationships.ToList();
            var hyperlinkIds = new HashSet<string>(hyperlinkRelationships.Select(relationship => relationship.Id), StringComparer.Ordinal);
            foreach (var hyperlink in worksheet.Elements<Hyperlinks>().SelectMany(links => links.Elements<Hyperlink>())) {
                string? relationshipId = hyperlink.Id?.Value;
                if (!string.IsNullOrEmpty(relationshipId) && !hyperlinkIds.Contains(relationshipId!)) {
                    skipReason = "Worksheet hyperlink relationships do not match hyperlink entries.";
                    return false;
                }
            }

            foreach (var tableDefinitionPart in tableDefinitionParts) {
                var table = tableDefinitionPart.Table;
                if (table == null
                    || table.Reference == null
                    || table.TableColumns == null
                    || table.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    skipReason = "Worksheet contains unsupported table metadata.";
                    return false;
                }
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return true;
            }

            foreach (var row in sheetData.Elements<Row>()) {
                if (!IsSimpleRow(row)) {
                    skipReason = "Worksheet contains row formatting outside the simple writer surface.";
                    return false;
                }

                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.InlineString != null) {
                        if (cell.InlineString.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                            skipReason = "Worksheet inline strings contain unknown Open XML elements.";
                            return false;
                        }
                    }

                    if (cell.CellFormula != null
                        && cell.CellFormula.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Worksheet contains formula metadata outside the simple writer surface.";
                        return false;
                    }

                    var dataType = cell.DataType?.Value;
                    if (dataType != null
                        && dataType != CellValues.Number
                        && dataType != CellValues.SharedString
                        && dataType != CellValues.InlineString
                        && dataType != CellValues.String
                        && dataType != CellValues.Boolean) {
                        skipReason = "Worksheet contains unsupported cell data type '" + dataType.Value.ToString() + "'.";
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsSimpleRow(Row row) {
            foreach (var attribute in row.GetAttributes()) {
                if (!string.Equals(attribute.LocalName, "r", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "hidden", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "ht", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "customHeight", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "outlineLevel", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "collapsed", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return row.CustomFormat?.Value != true && row.StyleIndex == null;
        }

        private static void WriteContentTypesEntry(ZipArchive archive, bool hasStyles, bool hasSharedStrings, int worksheetCount, int tableCount) {
            var builder = new System.Text.StringBuilder(512 + worksheetCount * 160 + tableCount * 160);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            builder.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            builder.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
            builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 1; i <= worksheetCount; i++) {
                builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                builder.Append(i.ToString(CultureInfo.InvariantCulture));
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }

            if (hasStyles) {
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            }

            if (hasSharedStrings) {
                builder.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            }

            for (int i = 1; i <= tableCount; i++) {
                builder.Append("<Override PartName=\"/xl/tables/table");
                builder.Append(i.ToString(CultureInfo.InvariantCulture));
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            }

            builder.Append("</Types>");
            WriteTextEntry(archive, "[Content_Types].xml", builder.ToString());
        }

        private static void WriteWorkbookEntry(ZipArchive archive, FastWorkbookPackageModel model) {
            var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
            var worksheets = model.Worksheets;
            using var stream = entry.Open();
            using var writer = CreateFastXmlWriter(stream);
            writer.WriteStartDocument();
            writer.WriteStartElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (model.FileVersion != null) {
                model.FileVersion.WriteTo(writer);
            }

            if (model.FileSharing != null) {
                model.FileSharing.WriteTo(writer);
            }

            if (model.WorkbookProperties != null) {
                model.WorkbookProperties.WriteTo(writer);
            }

            if (model.WorkbookProtection != null) {
                model.WorkbookProtection.WriteTo(writer);
            }

            if (model.BookViews != null) {
                model.BookViews.WriteTo(writer);
            }

            writer.WriteStartElement("sheets");
            foreach (var worksheet in worksheets) {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", worksheet.SheetName);
                writer.WriteAttributeString("sheetId", worksheet.SheetId.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", worksheet.WorkbookRelationshipId);
                if (!string.IsNullOrEmpty(worksheet.SheetState)) {
                    writer.WriteAttributeString("state", worksheet.SheetState);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            if (model.DefinedNames != null) {
                model.DefinedNames.WriteTo(writer);
            }

            if (model.CalculationProperties != null) {
                model.CalculationProperties.WriteTo(writer);
            }

            writer.WriteEndElement();
        }

        private static void WriteWorkbookRelationshipsEntry(ZipArchive archive, IReadOnlyList<FastWorksheetPackageModel> worksheets, bool hasStyles, bool hasSharedStrings) {
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

            if (hasSharedStrings) {
                builder.Append("<Relationship Id=\"rId");
                builder.Append((worksheets.Count + (hasStyles ? 2 : 1)).ToString(CultureInfo.InvariantCulture));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", builder.ToString());
        }

        private static void WriteCorePropertiesEntry(ZipArchive archive) {
            WriteTextEntry(archive, "docProps/core.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
                "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>");
        }

        private static void WriteAppPropertiesEntry(ZipArchive archive) {
            WriteTextEntry(archive, "docProps/app.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>OfficeIMO.Excel</Application>" +
                "</Properties>");
        }

        private static void WriteSharedStringsEntry(ZipArchive archive, SharedStringTable sharedStrings) {
            WriteTextEntry(
                archive,
                "xl/sharedStrings.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + sharedStrings.OuterXml);
        }

        private static void WriteWorksheetRelationshipsEntry(ZipArchive archive, FastWorksheetPackageModel worksheet) {
            var builder = new System.Text.StringBuilder(160 + worksheet.TablePartPaths.Count * 180 + worksheet.HyperlinkRelationships.Count * 220);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var item in worksheet.TablePartPaths.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, item.Key);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"");
                AppendXmlEscaped(builder, item.Value);
                builder.Append("\"/>");
            }

            foreach (var relationship in worksheet.HyperlinkRelationships.OrderBy(static item => item.Id, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, relationship.Id);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"");
                AppendXmlEscaped(builder, relationship.Target);
                builder.Append('"');
                if (relationship.IsExternal) {
                    builder.Append(" TargetMode=\"External\"");
                }

                builder.Append("/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, worksheet.RelationshipsPath, builder.ToString());
        }

        private static void WriteWorksheetEntry(ZipArchive archive, FastWorksheetPackageModel model) {
            var entry = archive.CreateEntry(model.WorksheetPath, CompressionLevel.Fastest);
            var worksheet = model.Worksheet;
            string dimension = worksheet.SheetDimension?.Reference?.Value ?? ExcelSheet.ComputeSheetDimensionReference(worksheet);
            var builder = new System.Text.StringBuilder(4096);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            if (model.HasRelationships) {
                builder.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
            }

            builder.Append(">");
            WriteBuilderAndClear(writer, builder);
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetProperties>());

            builder.Append("<dimension ref=\"");
            AppendXmlEscaped(builder, dimension);
            builder.Append("\"/>");
            WriteBuilderAndClear(writer, builder);

            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetViews>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetFormatProperties>());

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                AppendColumns(builder, columns);
                WriteBuilderAndClear(writer, builder);
            }

            writer.Write("<sheetData>");

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                foreach (var row in sheetData.Elements<Row>()) {
                    AppendSimpleRowStart(builder, row);

                    foreach (var cell in row.Elements<Cell>()) {
                        AppendSimpleCell(builder, cell);
                    }

                    builder.Append("</row>");
                    WriteBuilderAndClear(writer, builder);
                }
            }

            writer.Write("</sheetData>");
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetCalculationProperties>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetProtection>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<Scenarios>());

            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                writer.Write(autoFilter.OuterXml);
            }

            WriteOptionalElement(writer, worksheet.GetFirstChild<SortState>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<MergeCells>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PhoneticProperties>());
            WriteOptionalElements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>(writer, worksheet);
            WriteOptionalElement(writer, worksheet.GetFirstChild<DataValidations>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<Hyperlinks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PrintOptions>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PageMargins>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PageSetup>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<HeaderFooter>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<RowBreaks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<ColumnBreaks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<CellWatches>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors>());

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
                WriteBuilderAndClear(writer, builder);
            }

            writer.Write("</worksheet>");
        }

        private static void WriteBuilderAndClear(StreamWriter writer, System.Text.StringBuilder builder) {
            if (builder.Length == 0) {
                return;
            }

#if NET6_0_OR_GREATER
            writer.Write(builder);
#else
            writer.Write(builder.ToString());
#endif
            builder.Clear();
            if (builder.Capacity > 65536) {
                builder.Capacity = 4096;
            }
        }

        private static void WriteOptionalElement(StreamWriter writer, OpenXmlElement? element) {
            if (element != null) {
                writer.Write(element.OuterXml);
            }
        }

        private static void WriteOptionalElements<TElement>(StreamWriter writer, OpenXmlElement parent)
            where TElement : OpenXmlElement {
            foreach (var element in parent.Elements<TElement>()) {
                writer.Write(element.OuterXml);
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
                AppendBooleanAttribute(builder, "phonetic", column.Phonetic);
                builder.Append("/>");
            }

            builder.Append("</cols>");
        }

        private static void AppendSimpleRowStart(System.Text.StringBuilder builder, Row row) {
            builder.Append("<row");
            AppendUIntAttribute(builder, "r", row.RowIndex);
            AppendBooleanAttribute(builder, "hidden", row.Hidden);
            if (row.Height != null) {
                builder.Append(" ht=\"");
                builder.Append(row.Height.Value.ToString(CultureInfo.InvariantCulture));
                builder.Append('"');
            }

            AppendBooleanAttribute(builder, "customHeight", row.CustomHeight);
            AppendByteAttribute(builder, "outlineLevel", row.OutlineLevel);
            AppendBooleanAttribute(builder, "collapsed", row.Collapsed);
            builder.Append('>');
        }

        private static void AppendSimpleCell(System.Text.StringBuilder builder, Cell cell) {
            string? text = cell.CellValue?.Text;
            var dataType = cell.DataType?.Value;

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

            if (dataType == CellValues.Number) {
                builder.Append(" t=\"n\"");
            } else if (dataType == CellValues.SharedString) {
                builder.Append(" t=\"s\"");
            } else if (dataType == CellValues.InlineString || cell.InlineString != null) {
                builder.Append(" t=\"inlineStr\"");
            } else if (dataType == CellValues.String) {
                builder.Append(" t=\"str\"");
            } else if (dataType == CellValues.Boolean) {
                builder.Append(" t=\"b\"");
            }

            builder.Append('>');
            if (cell.CellFormula != null) {
                builder.Append(cell.CellFormula.OuterXml);
            }

            if (cell.InlineString != null) {
                builder.Append(cell.InlineString.OuterXml);
                builder.Append("</c>");
                return;
            }

            if (cell.CellValue != null) {
                builder.Append("<v>");
                AppendXmlEscaped(builder, text ?? string.Empty);
                builder.Append("</v>");
            }

            builder.Append("</c>");
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
