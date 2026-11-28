using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string PivotTableRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "pivotTable";

        /// <summary>
        /// Names the pivot tables whose location holds no cell values. Excel stores a
        /// pivot's rendered output as ordinary cells; a pivot written without them only
        /// appears after a spreadsheet application refreshes it.
        /// </summary>
        internal IReadOnlyList<string> GetPivotTablesWithoutRenderedCells() {
            MaterializePendingDirectCellValues();
            return ReadWorksheetObject(worksheet => {
                if (DeferredEntry is { } entry &&
                    _excelDocument.TryReadWorksheetPackage(entry, ReadPivotDefinitions, out List<(string Name, string Location)> definitions)) {
                    return FindUnrenderedPivotTables(worksheet, definitions);
                }

                return FindUnrenderedPivotTables(worksheet, ExcelPartRef.WorksheetForMetadata(this).PartsOfKind(ExcelPartKind.Pivot)
                    .Select(part => ReadPivotLocation(part.Root)));
            });
        }

        private static List<(string Name, string Location)> ReadPivotDefinitions(OfficeOpcPackage package, ExcelDeferredSheet sheet) {
            var definitions = new List<(string Name, string Location)>();
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(sheet.PartName)) {
                if (relationship.Type != PivotTableRelationship || relationship.TargetPartName is not string name || !package.PartExists(name)) continue;
                using OfficePooledBytes bytes = package.ReadPart(name);
                OfficeXmlTree tree = OfficeXmlTree.Parse(bytes.Array, bytes.Length, package.Namespaces);
                definitions.Add(ReadPivotLocation(new OfficeXmlNode(tree, tree.Root)));
            }
            return definitions;
        }

        private static (string Name, string Location) ReadPivotLocation(OfficeXmlNode definition) =>
            (definition.Attribute("name") ?? string.Empty, definition.Child(WorksheetMarkup, "location").Attribute("ref") ?? string.Empty);

        private static IReadOnlyList<string> FindUnrenderedPivotTables(OfficeXmlNode worksheet, IEnumerable<(string Name, string Location)> definitions) {
            var unrendered = new List<string>();
            List<(int Row, int Column)>? values = null;
            foreach ((string name, string location) in definitions) {
                if (name.Length == 0 || !ExcelA1.TryParseRange(location, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) continue;
                values ??= ReadPivotOutputCells(worksheet);
                if (!values.Any(value => value.Row >= firstRow && value.Row <= lastRow &&
                    value.Column >= firstColumn && value.Column <= lastColumn)) unrendered.Add(name);
            }
            return unrendered;
        }

        private static List<(int Row, int Column)> ReadPivotOutputCells(OfficeXmlNode worksheet) {
            var cells = new List<(int Row, int Column)>();
            int rowIndex = 0;
            foreach (OfficeXmlNode row in worksheet.Child(WorksheetMarkup, "sheetData").Children(WorksheetMarkup, "row")) {
                if (row.UInt32Attribute("r") is uint explicitRow && explicitRow > 0) {
                    rowIndex = checked((int)explicitRow);
                } else {
                    int referencedRow = 0;
                    foreach (OfficeXmlNode cell in row.Children(WorksheetMarkup, "c")) {
                        if (ExcelA1.TryParseCellReferenceFast(cell.Attribute("r"), out referencedRow, out _)) break;
                    }
                    rowIndex = referencedRow > 0 ? referencedRow : rowIndex + 1;
                }

                int nextColumn = 1;
                foreach (OfficeXmlNode cell in row.Children(WorksheetMarkup, "c")) {
                    string? reference = cell.Attribute("r");
                    int column = string.IsNullOrEmpty(reference) ? 0 : ExcelA1.ParseColumnIndexFromCellReferenceFast(reference!);
                    if (column <= 0) column = nextColumn;
                    nextColumn = column + 1;
                    if (cell.Child(WorksheetMarkup, "v").Exists || cell.Child(WorksheetMarkup, "is").Exists) cells.Add((rowIndex, column));
                }
            }
            return cells;
        }
    }
}
