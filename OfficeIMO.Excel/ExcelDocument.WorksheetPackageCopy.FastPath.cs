using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool TryCopyWorksheetPartGraphToEmptyWorkbook(
            ExcelDocument sourceDocument,
            ExcelSheet sourceSheet,
            string validatedName,
            out WorksheetPackageCopyResult? result) {
            result = null;
            WorksheetPart sourcePart = sourceSheet.WorksheetPart;
            if (!CanAdoptSourceWorkbookIndexParts(sourceDocument)
                || !CanCopyWorksheetPartGraphDirectly(sourceDocument, sourcePart)) {
                return false;
            }

            AdoptSourceWorkbookIndexParts(sourceDocument);
            WorksheetPart copiedPart = WorkbookPartRoot.AddPart(sourcePart);
            IReadOnlyDictionary<string, string> tableNameMap = BuildIdentityTableNameMap(copiedPart);
            Sheet sheet = AppendWorksheetElement(copiedPart, validatedName);
            var targetSheet = new ExcelSheet(this, _spreadSheetDocument, sheet);
            _tableNameCache = null;
            _nextTableId = null;
            MarkSheetCacheDirty();
            MarkPackageDirty();
            _requiresSavePreflight = false;
            WorkbookRoot.Save();
            result = new WorksheetPackageCopyResult(targetSheet, tableNameMap);
            return true;
        }

        private bool CanAdoptSourceWorkbookIndexParts(ExcelDocument sourceDocument) {
            WorkbookPart sourceWorkbookPart = sourceDocument.WorkbookPartRoot;
            if (DateSystem != sourceDocument.DateSystem
                || ReadSheetElements().Any()
                || WorkbookPartRoot.WorksheetParts.Any()
                || WorkbookPartRoot.WorkbookStylesPart != null
                || WorkbookPartRoot.SharedStringTablePart != null
                || WorkbookPartRoot.GetPartsOfType<ThemePart>().Any()
                || sourceWorkbookPart.WorkbookStylesPart?.IsRootElementLoaded == true
                || sourceWorkbookPart.SharedStringTablePart?.IsRootElementLoaded == true
                || sourceWorkbookPart.GetPartsOfType<ThemePart>().Any(static part => part.IsRootElementLoaded)) {
                return false;
            }

            return true;
        }

        private static bool CanCopyWorksheetPartGraphDirectly(ExcelDocument sourceDocument, WorksheetPart sourcePart) {
            WorkbookPart sourceWorkbookPart = sourceDocument.WorkbookPartRoot;
            if (sourcePart.IsRootElementLoaded
                || sourceWorkbookPart.SharedStringTablePart != null
                || sourceWorkbookPart.CellMetadataPart != null
                || sourceDocument.WorkbookRoot.ExternalReferences != null
                || sourceDocument.WorkbookRoot.DefinedNames?.Elements<DefinedName>().Any() == true
                || sourcePart.ExternalRelationships.Any()
                || sourcePart.HyperlinkRelationships.Any()
                || sourcePart.Parts.Any(pair => pair.OpenXmlPart is not TableDefinitionPart)
                || PartContainsFormulaOrCellMetadata(sourcePart)) {
                return false;
            }

            return sourcePart.TableDefinitionParts.All(tablePart =>
                !tablePart.IsRootElementLoaded
                && !tablePart.Parts.Any()
                && !tablePart.ExternalRelationships.Any()
                && !PartContainsFormulaOrCellMetadata(tablePart));
        }

        private void AdoptSourceWorkbookIndexParts(ExcelDocument sourceDocument) {
            WorkbookPart sourceWorkbookPart = sourceDocument.WorkbookPartRoot;
            if (sourceWorkbookPart.WorkbookStylesPart is WorkbookStylesPart sourceStylesPart) {
                WorkbookPartRoot.AddPart(sourceStylesPart);
            }

            ThemePart? sourceThemePart = sourceWorkbookPart.GetPartsOfType<ThemePart>().FirstOrDefault();
            if (sourceThemePart != null) {
                WorkbookPartRoot.AddPart(sourceThemePart);
            }
        }

        private static IReadOnlyDictionary<string, string> BuildIdentityTableNameMap(WorksheetPart copiedPart) {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (TableDefinitionPart tablePart in copiedPart.TableDefinitionParts) {
                string? name = tablePart.Table?.Name?.Value;
                string? displayName = tablePart.Table?.DisplayName?.Value;
                if (!string.IsNullOrWhiteSpace(name)) {
                    map[name!] = name!;
                }

                if (!string.IsNullOrWhiteSpace(displayName)) {
                    map[displayName!] = displayName!;
                }
            }

            return map;
        }

        private static bool PartContainsFormulaOrCellMetadata(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                XmlResolver = null
            });
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element) {
                    continue;
                }

                string localName = reader.LocalName;
                if (string.Equals(localName, "f", StringComparison.OrdinalIgnoreCase)
                    || localName.StartsWith("formula", StringComparison.OrdinalIgnoreCase)
                    || localName.StartsWith("queryTable", StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }

                if (!reader.HasAttributes) {
                    continue;
                }

                bool isCell = string.Equals(localName, "c", StringComparison.OrdinalIgnoreCase);
                for (bool hasAttribute = reader.MoveToFirstAttribute(); hasAttribute; hasAttribute = reader.MoveToNextAttribute()) {
                    if ((isCell && (string.Equals(reader.LocalName, "cm", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(reader.LocalName, "vm", StringComparison.OrdinalIgnoreCase)))
                        || string.Equals(reader.LocalName, "connectionId", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(reader.LocalName, "queryTableFieldId", StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }

                reader.MoveToElement();
            }

            return false;
        }
    }
}
