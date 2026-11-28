using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Reflects native record changes back into the editable worksheet after a committed save.</summary>
    internal static class XlsbWorksheetIdentity {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;

        internal static void Restore(OfficeXmlNode worksheet, XlsbSheet native) {
            if (native.UsedRange != null) {
                OfficeXmlNode dimension = XlsbWorksheetMarkup.Child(worksheet, "dimension");
                if (dimension.IsNone) {
                    OfficeXmlNode properties = XlsbWorksheetMarkup.Child(worksheet, "sheetPr");
                    dimension = worksheet.InsertBefore(XlsbWorksheetMarkup.Create(worksheet, "dimension"),
                        properties.Exists ? properties.NextSibling : worksheet.FirstChild);
                }
                dimension.SetAttribute("ref", native.UsedRange.ToA1Reference());
            }
            RestoreFormulas(worksheet, native.Cells);
            if (!string.IsNullOrEmpty(native.Properties?.CodeName)) {
                OfficeXmlNode properties = XlsbWorksheetMarkup.Child(worksheet, "sheetPr");
                if (properties.IsNone)
                    properties = worksheet.InsertBefore(XlsbWorksheetMarkup.Create(worksheet, "sheetPr"), worksheet.FirstChild);
                properties.SetAttribute("codeName", native.Properties!.CodeName);
            }
        }

        private static void RestoreFormulas(OfficeXmlNode worksheet, IReadOnlyList<XlsbCell> nativeCells) {
            Dictionary<string, XlsbCell>? changes = null;
            foreach (XlsbCell cell in nativeCells) {
                if (cell.FormulaText != null
                    || cell.FormulaBytes != null && cell.Kind == XlsbCellValueKind.Text && cell.Value is "") {
                    (changes ??= new(StringComparer.Ordinal)).Add(ExcelA1.CellReference(cell.Row, cell.Column), cell);
                }
            }
            if (changes == null) return;

            OfficeXmlNode sheetData = XlsbWorksheetMarkup.Child(worksheet, "sheetData");
            foreach (OfficeXmlNode row in sheetData.Children(Spreadsheet, "row")) {
                foreach (OfficeXmlNode cell in row.Children(Spreadsheet, "c")) {
                    if (cell.Attribute("r") is not string reference || !changes.TryGetValue(reference, out XlsbCell? native)) continue;
                    OfficeXmlNode formula = cell.Child(Spreadsheet, "f");
                    if (formula.Exists) {
                        if (native.FormulaText != null) formula.SetInnerText(native.FormulaText);
                    } else if (native.FormulaText == null) {
                        cell.Child(Spreadsheet, "is").Remove();
                        cell.SetAttribute("t", "str");
                        OfficeXmlNode value = cell.Child(Spreadsheet, "v");
                        if (value.IsNone) value = XlsbWorksheetMarkup.Append(cell, "v");
                        value.SetInnerText(string.Empty);
                    }
                }
            }
        }
    }
}
