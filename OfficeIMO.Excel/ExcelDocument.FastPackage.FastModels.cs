using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {

        private static class FastWorkbookPackageWriter {
            internal static void Write(Stream destination, FastWorkbookPackageModel model, CancellationToken ct) {
                using (var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true)) {
                    ct.ThrowIfCancellationRequested();
                    WriteContentTypesEntry(archive, model.HasStyles, model.HasSharedStrings, model.Worksheets.Count, model.Tables.Count);
                    ct.ThrowIfCancellationRequested();
                    WriteTextEntry(archive, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                        "</Relationships>");
                    WriteCorePropertiesEntry(archive);
                    WriteAppPropertiesEntry(archive);
                    ct.ThrowIfCancellationRequested();
                    WriteWorkbookEntry(archive, model);
                    WriteWorkbookRelationshipsEntry(archive, model.Worksheets, model.HasStyles, model.HasSharedStrings);
                    if (model.Stylesheet != null) {
                        WriteTextEntry(archive, "xl/styles.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + model.Stylesheet.OuterXml);
                    }

                    if (model.HasSharedStrings && model.SharedStrings != null) {
                        WriteSharedStringsEntry(archive, model.SharedStrings);
                    }

                    foreach (var worksheet in model.Worksheets) {
                        ct.ThrowIfCancellationRequested();
                        WriteWorksheetEntry(archive, worksheet);
                        if (worksheet.HasRelationships) {
                            WriteWorksheetRelationshipsEntry(archive, worksheet);
                        }
                    }

                    for (int i = 0; i < model.Tables.Count; i++) {
                        ct.ThrowIfCancellationRequested();
                        WriteTextEntry(
                            archive,
                            "xl/tables/table" + InvariantNumberText.Get(i + 1) + ".xml",
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
                        tablePartPaths[relId] = "../tables/table" + InvariantNumberText.Get(tableIndex) + ".xml";
                        tableIndex++;
                    }

                    var hyperlinkRelationships = worksheetPart.HyperlinkRelationships
                        .Select(relationship => new FastHyperlinkRelationshipModel(
                            relationship.Id,
                            relationship.Uri.ToString(),
                            relationship.IsExternal))
                        .ToList();

                    worksheets.Add(new FastWorksheetPackageModel(
                        sheet.Name?.Value ?? "Sheet" + InvariantNumberText.Get(sheetIndex + 1),
                        sheet.SheetId?.Value ?? (uint)(sheetIndex + 1),
                        GetSheetStateText(sheet),
                        "rId" + InvariantNumberText.Get(sheetIndex + 1),
                        "xl/worksheets/sheet" + InvariantNumberText.Get(sheetIndex + 1) + ".xml",
                        "xl/worksheets/_rels/sheet" + InvariantNumberText.Get(sheetIndex + 1) + ".xml.rels",
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

            internal bool RequiresRelationshipNamespace
                => HasRelationships || Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>()?.Id != null;
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
    }
}
