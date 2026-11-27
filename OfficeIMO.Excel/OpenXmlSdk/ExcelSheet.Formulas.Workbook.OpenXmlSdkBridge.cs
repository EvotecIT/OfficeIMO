using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>The workbook's sheet entries in order, each marked as a worksheet or not.</summary>
        internal List<ExcelFormulaSheetEntry> ReadFormulaSheetEntries() {
            var worksheetRelationshipIds = new HashSet<string>(
                WorkbookPartRoot.Parts
                    .Where(pair => pair.OpenXmlPart is WorksheetPart)
                    .Select(pair => pair.RelationshipId),
                StringComparer.Ordinal);
            var entries = new List<ExcelFormulaSheetEntry>();
            foreach (Sheet sheet in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                string? relationshipId = sheet.Id?.Value;
                entries.Add(new ExcelFormulaSheetEntry(sheet.Name?.Value,
                    !string.IsNullOrWhiteSpace(relationshipId) && worksheetRelationshipIds.Contains(relationshipId!), sheet));
            }
            return entries;
        }

        /// <summary>The workbook's sheet names in order.</summary>
        internal List<string?> ReadFormulaSheetNames() =>
            (WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()).Select(sheet => sheet.Name?.Value).ToList();

        /// <summary>The workbook's defined names in order.</summary>
        internal List<ExcelFormulaDefinedName> ReadFormulaDefinedNames() {
            var names = new List<ExcelFormulaDefinedName>();
            foreach (DefinedName definedName in WorkbookRoot.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>())
                names.Add(new ExcelFormulaDefinedName(definedName.Name?.Value, definedName.Text, definedName.LocalSheetId?.Value));
            return names;
        }

        /// <summary>A sheet wrapper for a workbook sheet entry, outside the document's sheet cache.</summary>
        private ExcelSheet CreateFormulaSheetWrapper(ExcelFormulaSheetEntry entry) =>
            new ExcelSheet(_excelDocument, _spreadSheetDocument, (Sheet)entry.Element);

        /// <summary>The workbook sheet entry with the name, if it has a relationship.</summary>
        private ExcelFormulaSheetEntry? FindFormulaSheetEntry(string sheetName) {
            Sheet? sheet = WorkbookRoot.Sheets?.Elements<Sheet>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            return sheet?.Id == null ? null : new ExcelFormulaSheetEntry(sheet.Name?.Value, true, sheet);
        }

        private static ExcelFormulaTable ToFormulaTable(Table table) => new ExcelFormulaTable(
            table.Name?.Value,
            table.DisplayName?.Value,
            table.Reference?.Value,
            table.HeaderRowCount?.Value ?? 1U,
            table.TotalsRowCount?.Value ?? 1U,
            table.TotalsRowShown?.Value == true,
            (table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()).Select(column => column.Name?.Value).ToList());

        /// <summary>The tables of this sheet.</summary>
        private IEnumerable<ExcelFormulaTable> ReadFormulaTables() =>
            _worksheetPart.TableDefinitionParts.Select(part => part.Table).OfType<Table>().Select(ToFormulaTable);

        /// <summary>The tables of every worksheet in the workbook.</summary>
        private IEnumerable<ExcelFormulaTable> ReadWorkbookFormulaTables() {
            WorkbookPart? workbookPart = _spreadSheetDocument.WorkbookPart;
            if (workbookPart == null) yield break;
            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                foreach (Table table in worksheetPart.TableDefinitionParts.Select(part => part.Table).OfType<Table>())
                    yield return ToFormulaTable(table);
        }

        /// <summary>Finds the table with the name or display name, and the sheet entry that holds it.</summary>
        private bool TryFindFormulaTable(string tableName, out ExcelFormulaSheetEntry sheetEntry, out ExcelFormulaTable table) {
            sheetEntry = null!;
            table = null!;
            WorkbookPart? workbookPart = _spreadSheetDocument.WorkbookPart;
            if (workbookPart == null) return false;
            foreach (var sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value == null) continue;
                if (workbookPart.GetPartById(sheetElement.Id.Value) is not WorksheetPart worksheetPart) continue;
                foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                    Table? candidate = tablePart.Table;
                    if (candidate == null
                        || (!string.Equals(candidate.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(candidate.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))) {
                        continue;
                    }
                    sheetEntry = new ExcelFormulaSheetEntry(sheetElement.Name?.Value, true, sheetElement);
                    table = ToFormulaTable(candidate);
                    return true;
                }
            }
            return false;
        }

        private bool TryGetContainingFormulaTable(int row, int column, out ExcelFormulaTable table) {
            table = null!;
            if (!HasTableDefinitions) return false;
            foreach (TableDefinitionPart part in _worksheetPart.TableDefinitionParts) {
                if (part.Table is not Table candidate) continue;
                if (candidate.Reference?.Value != null
                    && ExcelA1.TryParseRange(
                        candidate.Reference.Value.Replace("$", string.Empty),
                        out int firstRow,
                        out int firstColumn,
                        out int lastRow,
                        out int lastColumn)
                    && row >= firstRow
                    && row <= lastRow
                    && column >= firstColumn
                    && column <= lastColumn) {
                    table = ToFormulaTable(candidate);
                    return true;
                }
            }

            table = null!;
            return false;
        }
    }
}
