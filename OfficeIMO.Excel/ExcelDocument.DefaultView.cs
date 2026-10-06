using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal static bool IsNeutralWorkbookViews(BookViews views) => views.GetAttributes().Count == 0
            && views.ChildElements.Count == 1 && views.FirstChild is WorkbookView view
            && view.GetAttributes().Count == 0 && !view.HasChildren;
    }
}
