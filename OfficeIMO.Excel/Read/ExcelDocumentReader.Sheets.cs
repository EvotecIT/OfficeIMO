using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelDocumentReader {
        private OfficeXmlNode.ChildList WorkbookSheets =>
            ReadWorkbookNode.Child(Spreadsheet, "sheets").Children(Spreadsheet, "sheet");

        private static string? WorksheetRelationshipId(OfficeXmlNode sheet) => sheet.Attribute(Relationships, "id");

        public IReadOnlyList<string> GetSheetNames() {
            var names = new List<string>();
            foreach (OfficeXmlNode sheet in WorkbookSheets) {
                if (TryGetWorksheetPart(WorksheetRelationshipId(sheet), out _, out _)) {
                    names.Add(sheet.Attribute("name")!);
                }
            }
            return names;
        }

        public ExcelSheetReader GetSheet(string name) {
            foreach (OfficeXmlNode sheet in WorkbookSheets) {
                string? candidate = sheet.Attribute("name");
                if (string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase)
                    && TryGetWorksheetPart(WorksheetRelationshipId(sheet), out string path, out ExcelPartRef? live)) {
                    return CreateSheetReader(candidate!, path, live);
                }
            }
            throw new KeyNotFoundException($"Sheet '{name}' not found.");
        }

        public ExcelSheetReader GetSheet(int index) {
            ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
            foreach (OfficeXmlNode sheet in WorkbookSheets) {
                if (TryGetWorksheetPart(WorksheetRelationshipId(sheet), out string path, out ExcelPartRef? live) && --index == 0) {
                    return CreateSheetReader(sheet.Attribute("name")!, path, live);
                }
            }
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        public int SheetCount {
            get {
                int count = 0;
                foreach (OfficeXmlNode sheet in WorkbookSheets) {
                    if (TryGetWorksheetPart(WorksheetRelationshipId(sheet), out _, out _)) count++;
                }
                return count;
            }
        }

        private bool TryGetWorksheetPart(string? relationshipId, out string partName, out ExcelPartRef? livePart) {
            partName = string.Empty;
            livePart = default;
            if (string.IsNullOrEmpty(relationshipId)) return false;
            if (_liveWorkbook is ExcelPartRef workbook) {
                ExcelPartRef part = workbook.GetPartById(relationshipId);
                livePart = part;
                partName = part.Name;
                return part.Kind == ExcelPartKind.Worksheet;
            }

            var relationship = _package!.FindWorkbookRelationship(relationshipId!);
            if (relationship.Type != ExcelReadWorkbookParts.WorksheetRelationship || relationship.Target is not string target
                || !_package.PartExists(target)) return false;
            _package.ValidateWorksheetContentType(target);
            partName = target;
            return true;
        }
    }
}
