using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Excel imports row heights differently when a worksheet has no sheetView.
        // Keep a neutral workbook view in newly generated worksheets and after unfreezing.
        internal static SheetViews CreateDefaultSheetViews() => new SheetViews(new SheetView { WorkbookViewId = 0U });

        /// <summary>The serialized form of <see cref="CreateDefaultSheetViews"/>, for writers that emit XML without the element model.</summary>
        internal const string DefaultSheetViewsXml = "<x:sheetViews xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetView workbookViewId=\"0\" /></x:sheetViews>";

        internal static bool IsNeutralSheetViews(SheetViews views) => IsNeutralSheetViews(OfficeXmlNode.FromSdk(views));

        private static bool IsNeutralSheetViews(OfficeXmlNode views) {
            OfficeXmlNode view = views.FirstChild;
            if (views.HasAttributes || !view.Is(WorksheetMarkup, "sheetView") || view.NextSibling.Exists
                || view.Attributes().Count != 1 || view.UInt32Attribute("workbookViewId") != 0U) return false;

            OfficeXmlNode selection = view.FirstChild;
            return selection.IsNone || selection.Is(WorksheetMarkup, "selection") && selection.NextSibling.IsNone
                && selection.Attributes().Count == 2 && selection.FirstChild.IsNone
                && selection.Attribute("activeCell") == "A1" && selection.Attribute("sqref") == "A1";
        }

        internal static void EnsureDefaultSheetView(Worksheet worksheet) {
            SheetViews? views = worksheet.GetFirstChild<SheetViews>();
            if (views == null) worksheet.AddChild(CreateDefaultSheetViews(), true);
            else if (!views.Elements<SheetView>().Any()) views.Append(new SheetView { WorkbookViewId = 0U });
        }
    }
}
