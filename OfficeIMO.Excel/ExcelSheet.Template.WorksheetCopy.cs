using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal ExcelSheet CopyTemplateWorksheet(string sheetName) {
            ExcelSheet target = _excelDocument.AddWorkSheet(sheetName, SheetNameValidationMode.Sanitize);
            Worksheet worksheet = (Worksheet)WorksheetRoot.CloneNode(true);
            RemoveRelationshipBackedTemplateCopyElements(worksheet);
            target.WorksheetRoot = worksheet;
            RewriteTemplateWorksheetLocalReferences(target.WorksheetRoot, Name, target.Name);
            CopyTemplateWorksheetTableParts(target);
            CopyTemplateWorksheetHyperlinks(target);
            CopyTemplateWorksheetDrawings(target);
            CopyTemplateWorksheetComments(target);
            CopyTemplateWorksheetScopedDefinedNames(target);
            var workbookDefinedNameMap = CopyTemplateWorksheetWorkbookScopedDefinedNames(target);
            RewriteTemplateWorkbookDefinedNameReferences(target, workbookDefinedNameMap);
            target.WorksheetRoot.Save();
            target._sheetDataCache = null;
            target._lastAccessedRow = null;
            target._lastAccessedRowIndex = 0;
            target._lastAccessedCell = null;
            target._lastAccessedCellRowIndex = 0;
            target._lastAccessedCellColumnIndex = 0;
            target.ClearHeaderCache();
            target.MarkRequiresSavePreparation();
            return target;
        }

        private static void RewriteTemplateWorksheetLocalReferences(Worksheet worksheet, string sourceSheetName, string targetSheetName) {
            if (string.Equals(sourceSheetName, targetSheetName, StringComparison.Ordinal)) {
                return;
            }

            foreach (var formula in worksheet.Descendants<CellFormula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<Formula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<Formula1>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<Formula2>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var hyperlink in worksheet.Descendants<Hyperlink>()) {
                string? location = hyperlink.Location?.Value;
                if (string.IsNullOrEmpty(location)) {
                    continue;
                }

                string updated = ExcelDocument.ReplaceSheetNameReferences(location!, sourceSheetName, targetSheetName);
                if (!string.Equals(updated, location, StringComparison.Ordinal)) {
                    hyperlink.Location = updated;
                }
            }
        }

        private static void RewriteTemplateTableLocalReferences(Table table, string sourceSheetName, string targetSheetName) {
            if (string.Equals(sourceSheetName, targetSheetName, StringComparison.Ordinal)) {
                return;
            }

            foreach (var formula in table.Descendants<CalculatedColumnFormula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in table.Descendants<TotalsRowFormula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }
        }

        private static void RewriteTemplateFormulaText(OpenXmlLeafTextElement formula, string sourceSheetName, string targetSheetName) {
            string? text = formula.Text;
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            string updated = ExcelDocument.ReplaceSheetNameReferences(text, sourceSheetName, targetSheetName);
            if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                formula.Text = updated;
            }
        }

        private void CopyTemplateWorksheetScopedDefinedNames(ExcelSheet target) {
            DefinedNames? definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return;
            }

            ushort sourceSheetPosition = GetTemplateSheetPositionIndex(Name);
            ushort targetSheetPosition = GetTemplateSheetPositionIndex(target.Name);
            var sourceNames = definedNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId != null && name.LocalSheetId.Value == sourceSheetPosition)
                .ToList();
            if (sourceNames.Count == 0) {
                return;
            }

            foreach (DefinedName sourceName in sourceNames) {
                string? definedName = sourceName.Name?.Value;
                if (string.IsNullOrWhiteSpace(definedName)) {
                    continue;
                }

                foreach (DefinedName existing in definedNames.Elements<DefinedName>()
                    .Where(name => name.LocalSheetId != null
                        && name.LocalSheetId.Value == targetSheetPosition
                        && string.Equals(name.Name?.Value, definedName, StringComparison.OrdinalIgnoreCase))
                    .ToList()) {
                    existing.Remove();
                }

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.LocalSheetId = targetSheetPosition;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ExcelDocument.ReplaceSheetNameReferences(clone.Text!, Name, target.Name);
                }

                definedNames.Append(clone);
            }

            WorkbookRoot.Save();
        }

        private Dictionary<string, string> CopyTemplateWorksheetWorkbookScopedDefinedNames(ExcelSheet target) {
            var copiedNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            DefinedNames? definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return copiedNames;
            }

            var existingNames = new HashSet<string>(definedNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId == null && !string.IsNullOrWhiteSpace(name.Name?.Value))
                .Select(name => name.Name!.Value!), StringComparer.OrdinalIgnoreCase);
            var sourceNames = definedNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId == null
                    && !IsTemplateBuiltInDefinedName(name.Name?.Value)
                    && !string.IsNullOrWhiteSpace(name.Name?.Value)
                    && DefinedNameReferencesTemplateSheet(name.Text, Name, target.Name))
                .ToList();
            if (sourceNames.Count == 0) {
                return copiedNames;
            }

            foreach (DefinedName sourceName in sourceNames) {
                string sourceDefinedName = sourceName.Name!.Value!;
                string targetDefinedName = CreateTemplateWorkbookDefinedName(sourceDefinedName, target.Name, existingNames);
                existingNames.Add(targetDefinedName);

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.Name = targetDefinedName;
                clone.LocalSheetId = null;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ExcelDocument.ReplaceSheetNameReferences(clone.Text!, Name, target.Name);
                }

                definedNames.Append(clone);
                copiedNames[sourceDefinedName] = targetDefinedName;
            }

            WorkbookRoot.Save();
            return copiedNames;
        }

        private static bool DefinedNameReferencesTemplateSheet(string? reference, string sourceSheetName, string targetSheetName) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            string updated = ExcelDocument.ReplaceSheetNameReferences(reference!, sourceSheetName, targetSheetName);
            return !string.Equals(updated, reference, StringComparison.Ordinal);
        }

        private static bool IsTemplateBuiltInDefinedName(string? name) {
            return !string.IsNullOrWhiteSpace(name)
                && name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
        }

        private static string CreateTemplateWorkbookDefinedName(string sourceDefinedName, string targetSheetName, ISet<string> existingNames) {
            string baseName = Regex.Replace(sourceDefinedName + "_" + targetSheetName, @"[^A-Za-z0-9_\.]", "_");
            if (string.IsNullOrWhiteSpace(baseName) || (!char.IsLetter(baseName[0]) && baseName[0] != '_' && baseName[0] != '\\')) {
                baseName = "_" + baseName;
            }

            if (baseName.Length > 240) {
                baseName = baseName.Substring(0, 240);
            }

            string candidate = baseName;
            int suffix = 2;
            while (existingNames.Contains(candidate)) {
                string suffixText = "_" + suffix.ToString(CultureInfo.InvariantCulture);
                int maxBaseLength = Math.Max(1, 255 - suffixText.Length);
                candidate = (baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) : baseName) + suffixText;
                suffix++;
            }

            return candidate;
        }

        private static void RewriteTemplateWorkbookDefinedNameReferences(ExcelSheet target, IReadOnlyDictionary<string, string> definedNameMap) {
            if (definedNameMap.Count == 0) {
                return;
            }

            RewriteDefinedNameReferences(target.WorksheetRoot, definedNameMap);

            foreach (TableDefinitionPart tablePart in target._worksheetPart.TableDefinitionParts) {
                if (tablePart.Table != null) {
                    RewriteDefinedNameReferences(tablePart.Table, definedNameMap);
                    tablePart.Table.Save();
                }
            }

            DrawingsPart? drawingsPart = target._worksheetPart.DrawingsPart;
            if (drawingsPart != null) {
                foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                    if (chartPart.ChartSpace != null) {
                        RewriteDefinedNameReferences(chartPart.ChartSpace, definedNameMap);
                        chartPart.ChartSpace.Save();
                    }
                }
            }
        }

        private static void RewriteDefinedNameReferences(OpenXmlElement root, IReadOnlyDictionary<string, string> definedNameMap) {
            foreach (var formula in root.Descendants<CellFormula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<Formula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<Formula1>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<Formula2>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }
        }

        private static void RewriteDefinedNameFormulaText(OpenXmlLeafTextElement formula, IReadOnlyDictionary<string, string> definedNameMap) {
            string? text = formula.Text;
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            string updated = text!;
            foreach (var pair in definedNameMap) {
                updated = Regex.Replace(
                    updated,
                    @"(?<![A-Za-z0-9_\.])" + Regex.Escape(pair.Key) + @"(?![A-Za-z0-9_\.])",
                    pair.Value,
                    RegexOptions.IgnoreCase,
                    TimeSpan.FromMilliseconds(100));
            }

            if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                formula.Text = updated;
            }
        }

        private ushort GetTemplateSheetPositionIndex(string sheetName) {
            var sheets = WorkbookRoot.Sheets?.OfType<Sheet>().ToList() ?? new List<Sheet>();
            for (ushort index = 0; index < sheets.Count; index++) {
                if (string.Equals(sheets[index].Name?.Value, sheetName, StringComparison.Ordinal)) {
                    return index;
                }
            }

            throw new ArgumentException($"Worksheet '{sheetName}' was not found.", nameof(sheetName));
        }

        private static void RemoveRelationshipBackedTemplateCopyElements(Worksheet worksheet) {
            foreach (var drawing in worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Drawing>().ToList()) {
                drawing.Remove();
            }

            foreach (var legacyDrawing in worksheet.Descendants<LegacyDrawing>().ToList()) {
                legacyDrawing.Remove();
            }

            foreach (var legacyHeaderFooterDrawing in worksheet.Descendants<LegacyDrawingHeaderFooter>().ToList()) {
                legacyHeaderFooterDrawing.Remove();
            }

            foreach (var picture in worksheet.Descendants<Picture>().ToList()) {
                picture.Remove();
            }

            foreach (var oleObjects in worksheet.Descendants<OleObjects>().ToList()) {
                oleObjects.Remove();
            }

            foreach (var controls in worksheet.Descendants<Controls>().ToList()) {
                controls.Remove();
            }

        }

        private void CopyTemplateWorksheetTableParts(ExcelSheet target) {
            TableParts? clonedTableParts = target.WorksheetRoot.GetFirstChild<TableParts>();
            if (clonedTableParts == null) {
                return;
            }

            foreach (TablePart tablePart in clonedTableParts.Elements<TablePart>().ToList()) {
                string? sourceRelationshipId = tablePart.Id?.Value;
                if (string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                    tablePart.Remove();
                    continue;
                }

                if (_worksheetPart.GetPartById(sourceRelationshipId!) is not TableDefinitionPart sourceTablePart
                    || sourceTablePart.Table == null) {
                    tablePart.Remove();
                    continue;
                }

                string targetRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
                TableDefinitionPart targetTablePart = target._worksheetPart.AddNewPart<TableDefinitionPart>(targetRelationshipId);
                Table clonedTable = (Table)sourceTablePart.Table.CloneNode(true);
                RewriteTemplateTableLocalReferences(clonedTable, Name, target.Name);
                clonedTable.Id = _excelDocument.AllocateTableId();

                string requestedName = clonedTable.Name?.Value ?? clonedTable.DisplayName?.Value ?? "Table";
                string tableName = EnsureValidUniqueTableName(requestedName, TableNameValidationMode.Sanitize);
                clonedTable.Name = tableName;
                clonedTable.DisplayName = tableName;
                _excelDocument.ReserveTableName(tableName);

                targetTablePart.Table = clonedTable;
                targetTablePart.Table.Save();
                tablePart.Id = targetRelationshipId;
            }

            uint count = (uint)clonedTableParts.Elements<TablePart>().Count();
            if (count == 0) {
                clonedTableParts.Remove();
            } else {
                clonedTableParts.Count = count;
            }
        }

        private static string GetUnusedRelationshipId(OpenXmlPartContainer partContainer) {
            var existing = new HashSet<string>(StringComparer.Ordinal);
            foreach (var part in partContainer.Parts) {
                if (!string.IsNullOrWhiteSpace(part.RelationshipId)) {
                    existing.Add(part.RelationshipId);
                }
            }

            foreach (var relationship in partContainer.ExternalRelationships) {
                if (!string.IsNullOrWhiteSpace(relationship.Id)) {
                    existing.Add(relationship.Id);
                }
            }

            if (partContainer is WorksheetPart worksheetPart) {
                foreach (var relationship in worksheetPart.HyperlinkRelationships) {
                    if (!string.IsNullOrWhiteSpace(relationship.Id)) {
                        existing.Add(relationship.Id);
                    }
                }
            }

            foreach (var relationship in partContainer.DataPartReferenceRelationships) {
                if (!string.IsNullOrWhiteSpace(relationship.Id)) {
                    existing.Add(relationship.Id);
                }
            }

            int index = 1;
            string id;
            do {
                id = "rId" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            } while (existing.Contains(id));

            return id;
        }

        private void CopyTemplateWorksheetHyperlinks(ExcelSheet target) {
            Hyperlinks? hyperlinks = target.WorksheetRoot.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) {
                return;
            }

            foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>().Where(hyperlink => hyperlink.Id != null).ToList()) {
                string? sourceRelationshipId = hyperlink.Id?.Value;
                HyperlinkRelationship? sourceRelationship = string.IsNullOrWhiteSpace(sourceRelationshipId)
                    ? null
                    : _worksheetPart.HyperlinkRelationships.FirstOrDefault(relationship =>
                        string.Equals(relationship.Id, sourceRelationshipId, StringComparison.OrdinalIgnoreCase));

                if (sourceRelationship == null) {
                    hyperlink.Remove();
                    continue;
                }

                string targetRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
                target._worksheetPart.AddHyperlinkRelationship(sourceRelationship.Uri, sourceRelationship.IsExternal, targetRelationshipId);
                hyperlink.Id = targetRelationshipId;
            }

            if (!hyperlinks.Elements<Hyperlink>().Any()) {
                hyperlinks.Remove();
            }
        }

        private void CopyTemplateWorksheetDrawings(ExcelSheet target) {
            DocumentFormat.OpenXml.Spreadsheet.Drawing? sourceDrawing = WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            if (sourceDrawing?.Id?.Value is not string sourceRelationshipId || string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                return;
            }

            OpenXmlPart? sourcePart;
            try {
                sourcePart = _worksheetPart.GetPartById(sourceRelationshipId);
            } catch {
                return;
            }

            if (sourcePart is not DrawingsPart sourceDrawingsPart || sourceDrawingsPart.WorksheetDrawing == null) {
                return;
            }

            string targetRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
            DrawingsPart targetDrawingsPart = target._worksheetPart.AddNewPart<DrawingsPart>(targetRelationshipId);
            CopyPartStream(sourceDrawingsPart, targetDrawingsPart);
            targetDrawingsPart.WorksheetDrawing ??= new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            CopyTemplateDrawingPartRelationships(sourceDrawingsPart, targetDrawingsPart, Name, target.Name);
            targetDrawingsPart.WorksheetDrawing.Save();

            var drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = targetRelationshipId };
            LegacyDrawing? legacyDrawing = target.WorksheetRoot.GetFirstChild<LegacyDrawing>();
            LegacyDrawingHeaderFooter? legacyHeaderFooter = target.WorksheetRoot.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacyDrawing != null) {
                target.WorksheetRoot.InsertBefore(drawing, legacyDrawing);
            } else if (legacyHeaderFooter != null) {
                target.WorksheetRoot.InsertBefore(drawing, legacyHeaderFooter);
            } else {
                target.WorksheetRoot.Append(drawing);
            }
        }

        private static void CopyTemplateDrawingPartRelationships(
            DrawingsPart sourceDrawingsPart,
            DrawingsPart targetDrawingsPart,
            string sourceSheetName,
            string targetSheetName) {
            foreach (var relationship in sourceDrawingsPart.Parts.ToList()) {
                if (relationship.OpenXmlPart is ChartPart sourceChartPart) {
                    string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                    ChartPart targetChartPart = targetDrawingsPart.AddNewPart<ChartPart>(targetRelationshipId);
                    if (sourceChartPart.ChartSpace != null) {
                        targetChartPart.ChartSpace = (DocumentFormat.OpenXml.Drawing.Charts.ChartSpace)sourceChartPart.ChartSpace.CloneNode(true);
                        RewriteChartSheetReferences(targetChartPart, sourceSheetName, targetSheetName);
                        targetChartPart.ChartSpace.Save();
                    }

                    foreach (var chartRelationship in sourceChartPart.Parts.ToList()) {
                        CopyTemplateKnownPartRelationship(chartRelationship.OpenXmlPart, targetChartPart, chartRelationship.RelationshipId);
                    }

                    RewriteDrawingRelationshipId(targetDrawingsPart.WorksheetDrawing!, relationship.RelationshipId, targetRelationshipId);
                    continue;
                }

                if (relationship.OpenXmlPart is ImagePart sourceImagePart) {
                    string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                    CopyTemplateImagePart(sourceImagePart, targetDrawingsPart, targetRelationshipId);
                    RewriteDrawingRelationshipId(targetDrawingsPart.WorksheetDrawing!, relationship.RelationshipId, targetRelationshipId);
                    continue;
                }

                if (IsTemplateDiagramPart(relationship.OpenXmlPart)) {
                    string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                    CopyTemplateKnownPartRelationship(relationship.OpenXmlPart, targetDrawingsPart, targetRelationshipId);
                    RewriteDrawingRelationshipId(targetDrawingsPart.WorksheetDrawing!, relationship.RelationshipId, targetRelationshipId);
                    continue;
                }

                CopyTemplateKnownPartRelationship(relationship.OpenXmlPart, targetDrawingsPart, relationship.RelationshipId);
            }

            CopyReferencedDrawingImages(sourceDrawingsPart, targetDrawingsPart);
        }

        private static void RewriteChartSheetReferences(ChartPart chartPart, string sourceSheetName, string targetSheetName) {
            if (chartPart.ChartSpace == null || string.Equals(sourceSheetName, targetSheetName, StringComparison.Ordinal)) {
                return;
            }

            foreach (var formula in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()) {
                string? text = formula.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                string updated = ExcelDocument.ReplaceSheetNameReferences(text, sourceSheetName, targetSheetName);
                if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                    formula.Text = updated;
                }
            }
        }

        private static void CopyTemplateKnownPartRelationship(OpenXmlPart sourcePart, OpenXmlPartContainer targetContainer, string sourceRelationshipId) {
            if (sourcePart is ImagePart sourceImagePart) {
                ImagePart? targetImagePart = targetContainer switch {
                    ChartPart chartPart => chartPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    ChartDrawingPart chartDrawingPart => chartDrawingPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    DiagramDataPart diagramDataPart => diagramDataPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    DiagramLayoutDefinitionPart diagramLayoutPart => diagramLayoutPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    DiagramPersistLayoutPart diagramPersistLayoutPart => diagramPersistLayoutPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    _ => null
                };
                if (targetImagePart == null) {
                    return;
                }

                CopyPartStream(sourceImagePart, targetImagePart);
                return;
            }

            if (sourcePart is ChartStylePart sourceChartStylePart && targetContainer is ChartPart targetChartPartForStyle) {
                ChartStylePart targetChartStylePart = targetChartPartForStyle.AddNewPart<ChartStylePart>(sourceRelationshipId);
                CopyPartStream(sourceChartStylePart, targetChartStylePart);
                return;
            }

            if (sourcePart is ChartColorStylePart sourceChartColorStylePart && targetContainer is ChartPart targetChartPartForColorStyle) {
                ChartColorStylePart targetChartColorStylePart = targetChartPartForColorStyle.AddNewPart<ChartColorStylePart>(sourceRelationshipId);
                CopyPartStream(sourceChartColorStylePart, targetChartColorStylePart);
                return;
            }

            if (sourcePart is ChartDrawingPart sourceChartDrawingPart && targetContainer is ChartPart targetChartPartForDrawing) {
                ChartDrawingPart targetChartDrawingPart = targetChartPartForDrawing.AddNewPart<ChartDrawingPart>(sourceRelationshipId);
                CopyPartStream(sourceChartDrawingPart, targetChartDrawingPart);
                CopyTemplateChildPartRelationships(sourceChartDrawingPart, targetChartDrawingPart);

                return;
            }

            if (sourcePart is DiagramColorsPart sourceDiagramColorsPart && targetContainer is DrawingsPart targetDrawingsPartForColors) {
                DiagramColorsPart targetDiagramColorsPart = targetDrawingsPartForColors.AddNewPart<DiagramColorsPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramColorsPart, targetDiagramColorsPart);
                CopyTemplateChildPartRelationships(sourceDiagramColorsPart, targetDiagramColorsPart);
                return;
            }

            if (sourcePart is DiagramDataPart sourceDiagramDataPart && targetContainer is DrawingsPart targetDrawingsPartForData) {
                DiagramDataPart targetDiagramDataPart = targetDrawingsPartForData.AddNewPart<DiagramDataPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramDataPart, targetDiagramDataPart);
                CopyTemplateChildPartRelationships(sourceDiagramDataPart, targetDiagramDataPart);
                return;
            }

            if (sourcePart is DiagramLayoutDefinitionPart sourceDiagramLayoutPart && targetContainer is DrawingsPart targetDrawingsPartForLayout) {
                DiagramLayoutDefinitionPart targetDiagramLayoutPart = targetDrawingsPartForLayout.AddNewPart<DiagramLayoutDefinitionPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramLayoutPart, targetDiagramLayoutPart);
                CopyTemplateChildPartRelationships(sourceDiagramLayoutPart, targetDiagramLayoutPart);
                return;
            }

            if (sourcePart is DiagramPersistLayoutPart sourceDiagramPersistLayoutPart && targetContainer is DrawingsPart targetDrawingsPartForPersistLayout) {
                DiagramPersistLayoutPart targetDiagramPersistLayoutPart = targetDrawingsPartForPersistLayout.AddNewPart<DiagramPersistLayoutPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramPersistLayoutPart, targetDiagramPersistLayoutPart);
                CopyTemplateChildPartRelationships(sourceDiagramPersistLayoutPart, targetDiagramPersistLayoutPart);
                return;
            }

            if (sourcePart is DiagramStylePart sourceDiagramStylePart && targetContainer is DrawingsPart targetDrawingsPartForStyle) {
                DiagramStylePart targetDiagramStylePart = targetDrawingsPartForStyle.AddNewPart<DiagramStylePart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramStylePart, targetDiagramStylePart);
                CopyTemplateChildPartRelationships(sourceDiagramStylePart, targetDiagramStylePart);
                return;
            }

            if (sourcePart is EmbeddedPackagePart sourceEmbeddedPackagePart && targetContainer is ChartPart targetChartPartForEmbeddedPackage) {
                EmbeddedPackagePart targetEmbeddedPackagePart = targetChartPartForEmbeddedPackage.AddEmbeddedPackagePart(sourceEmbeddedPackagePart.ContentType, sourceRelationshipId);
                CopyPartStream(sourceEmbeddedPackagePart, targetEmbeddedPackagePart);
            }
        }

        private static bool IsTemplateDiagramPart(OpenXmlPart sourcePart) {
            return sourcePart is DiagramColorsPart
                || sourcePart is DiagramDataPart
                || sourcePart is DiagramLayoutDefinitionPart
                || sourcePart is DiagramPersistLayoutPart
                || sourcePart is DiagramStylePart;
        }

        private static void CopyTemplateChildPartRelationships(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            foreach (var relationship in sourcePart.Parts.ToList()) {
                CopyTemplateKnownPartRelationship(relationship.OpenXmlPart, targetPart, relationship.RelationshipId);
            }
        }

        private static void CopyTemplateImagePart(ImagePart sourceImagePart, DrawingsPart targetDrawingsPart, string targetRelationshipId) {
            ImagePart targetImagePart = targetDrawingsPart.AddImagePart(sourceImagePart.ContentType, targetRelationshipId);
            CopyPartStream(sourceImagePart, targetImagePart);
        }

        private static void CopyReferencedDrawingImages(DrawingsPart sourceDrawingsPart, DrawingsPart targetDrawingsPart) {
            if (targetDrawingsPart.WorksheetDrawing == null) {
                return;
            }

            foreach (var blip in targetDrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()) {
                string? sourceRelationshipId = blip.Embed?.Value;
                if (string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                    continue;
                }

                if (TryGetPartById(targetDrawingsPart, sourceRelationshipId!) != null) {
                    continue;
                }

                if (TryGetPartById(sourceDrawingsPart, sourceRelationshipId!) is not ImagePart sourceImagePart) {
                    continue;
                }

                string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                CopyTemplateImagePart(sourceImagePart, targetDrawingsPart, targetRelationshipId);
                blip.Embed = targetRelationshipId;
            }
        }

        private static OpenXmlPart? TryGetPartById(OpenXmlPartContainer container, string relationshipId) {
            try {
                return container.GetPartById(relationshipId);
            } catch {
                return null;
            }
        }

        private static void CopyPartStream(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            using (Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read))
            using (Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write)) {
                sourceStream.CopyTo(targetStream);
            }
        }

        private static void RewriteDrawingRelationshipId(OpenXmlElement root, string oldRelationshipId, string newRelationshipId) {
            foreach (var element in root.Descendants<OpenXmlElement>()) {
                foreach (var attribute in element.GetAttributes()) {
                    if (string.Equals(attribute.NamespaceUri, "http://schemas.openxmlformats.org/officeDocument/2006/relationships", StringComparison.Ordinal)
                        && string.Equals(attribute.Value, oldRelationshipId, StringComparison.Ordinal)) {
                        element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newRelationshipId));
                    }
                }
            }
        }

        private void CopyTemplateWorksheetComments(ExcelSheet target) {
            WorksheetCommentsPart? sourceCommentsPart = _worksheetPart.WorksheetCommentsPart;
            if (sourceCommentsPart?.Comments?.CommentList?.Elements<Comment>().Any() != true) {
                return;
            }

            string commentsRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
            WorksheetCommentsPart targetCommentsPart = target._worksheetPart.AddNewPart<WorksheetCommentsPart>(commentsRelationshipId);
            targetCommentsPart.Comments = (Comments)sourceCommentsPart.Comments.CloneNode(true);
            targetCommentsPart.Comments.Save();

            LegacyDrawing? sourceLegacyDrawing = WorksheetRoot.GetFirstChild<LegacyDrawing>();
            if (sourceLegacyDrawing?.Id?.Value is not string sourceLegacyRelationshipId
                || string.IsNullOrWhiteSpace(sourceLegacyRelationshipId)) {
                return;
            }

            OpenXmlPart? sourceLegacyPart;
            try {
                sourceLegacyPart = _worksheetPart.GetPartById(sourceLegacyRelationshipId);
            } catch {
                return;
            }

            if (sourceLegacyPart is not VmlDrawingPart sourceVmlPart) {
                return;
            }

            string vmlRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
            VmlDrawingPart targetVmlPart = target._worksheetPart.AddNewPart<VmlDrawingPart>(vmlRelationshipId);
            using (Stream sourceStream = sourceVmlPart.GetStream(FileMode.Open, FileAccess.Read))
            using (Stream targetStream = targetVmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                sourceStream.CopyTo(targetStream);
            }

            var legacyDrawing = new LegacyDrawing { Id = vmlRelationshipId };
            LegacyDrawingHeaderFooter? legacyHeaderFooter = target.WorksheetRoot.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacyHeaderFooter != null) {
                target.WorksheetRoot.InsertBefore(legacyDrawing, legacyHeaderFooter);
            } else {
                target.WorksheetRoot.Append(legacyDrawing);
            }
        }
    }
}
