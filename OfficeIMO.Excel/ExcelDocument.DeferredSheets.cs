using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>Worksheets added to a loaded workbook that has not opened its SDK package.</summary>
    public partial class ExcelDocument {
        private const string SpreadsheetMainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string NewWorksheetXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:worksheet xmlns:x=\"" + SpreadsheetMainNamespace
            + "\"><x:sheetViews><x:sheetView workbookViewId=\"0\" /></x:sheetViews><x:sheetData /></x:worksheet>";

        /// <summary>
        /// Adds a worksheet to the package of a workbook that has not opened its SDK package, as the model adds one: a
        /// part holding the default sheet view and an empty grid, its relationship, and the workbook's sheet entry.
        /// Returns null, having changed nothing, when the workbook part needs the model.
        /// </summary>
        private ExcelSheet? TryAddDeferredSheet(string name) {
            ExcelDeferredWorkbook workbook = _deferredWorkbook!;
            if (name.Length == 0) name = "Sheet1";
            uint sheetId = 1U;
            foreach (var existing in id) sheetId = Math.Max(sheetId, existing.Value + 1U);
            byte[] bytes;
            string partName, relationshipId;
            using (var input = new MemoryStream(_deferredPackageBytes!, writable: false))
            using (OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true)) {
                OfficeXmlTree tree = package.GetXml(workbook.WorkbookPartName);
                int sheets = tree.FindChild(tree.Root, OfficeXmlNamespaces.SpreadsheetML, "sheets");
                if (sheets == OfficeXmlTree.NoNode) return null;
                int number = workbook.Sheets.Count + 1;
                string directory = workbook.WorkbookPartName.Substring(0, workbook.WorkbookPartName.LastIndexOf('/') + 1);
                while (package.PartExists(directory + "worksheets/sheet" + number.ToString(CultureInfo.InvariantCulture) + ".xml")) number++;
                partName = directory + "worksheets/sheet" + number.ToString(CultureInfo.InvariantCulture) + ".xml";
                package.CreateBinaryPart(partName, WorksheetContentType, Encoding.UTF8.GetBytes(NewWorksheetXml));
                relationshipId = package.AddRelationship(workbook.WorkbookPartName, WorksheetRelationshipType, partName.Substring(directory.Length));
                int sheet = tree.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "sheet");
                tree.SetAttribute(sheet, OfficeXmlNamespaces.None, "name", name);
                tree.SetAttribute(sheet, OfficeXmlNamespaces.None, "sheetId", sheetId.ToString(CultureInfo.InvariantCulture));
                tree.SetAttribute(sheet, OfficeXmlNamespaces.Relationships, "id", relationshipId);
                tree.Append(sheets, sheet);
                using var output = new MemoryStream(_deferredPackageBytes!.Length + 1024);
                package.Save(output);
                bytes = output.ToArray();
            }

            var entry = new ExcelDeferredSheet(name, sheetId, relationshipId, partName, isWorksheet: true);
            workbook.Sheets.Add(entry);
            _deferredPackageBytes = bytes;
            id.Add(sheetId);
            MarkPackageDirty();
            var created = new ExcelSheet(this, entry);
            if (SheetCachingEnabled && !_sheetCacheDirty && _cachedSheets != null) {
                // The cache holds its own handle, as a rebuild would; both share the sheet's recorded changes.
                _cachedSheets.Add(new ExcelSheet(this, entry));
            } else {
                _sheetCacheDirty = true;
            }
            return created;
        }

        /// <summary>The names of the workbook's sheets, read without the SDK while its package is not open.</summary>
        private IEnumerable<string> DeferredSheetNames() {
            foreach (ExcelDeferredSheet entry in _deferredWorkbook!.Sheets) yield return entry.Name;
        }
    }
}
