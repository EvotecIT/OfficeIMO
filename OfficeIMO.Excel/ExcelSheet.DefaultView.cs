using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Excel imports row heights differently when a worksheet has no sheetView.
        // Keep a neutral workbook view in newly generated worksheets and after unfreezing.
        internal static SheetViews CreateDefaultSheetViews() => new SheetViews(new SheetView { WorkbookViewId = 0U });

        internal static bool IsNeutralSheetViews(SheetViews views) {
            if (views.GetAttributes().Count != 0 || views.ChildElements.Count != 1 || views.FirstChild is not SheetView view
                || view.GetAttributes().Count != 1 || view.WorkbookViewId?.Value != 0U) return false;
            if (!view.HasChildren) return true;
            return view.ChildElements.Count == 1 && view.FirstChild is Selection selection
                && selection.GetAttributes().Count == 2 && !selection.HasChildren
                && selection.ActiveCell?.Value == "A1" && selection.SequenceOfReferences?.InnerText == "A1";
        }

        internal static void EnsureDefaultSheetView(Worksheet worksheet) {
            SheetViews? views = worksheet.GetFirstChild<SheetViews>();
            if (views == null) worksheet.AddChild(CreateDefaultSheetViews(), true);
            else if (!views.Elements<SheetView>().Any()) views.Append(new SheetView { WorkbookViewId = 0U });
        }
    }
}
