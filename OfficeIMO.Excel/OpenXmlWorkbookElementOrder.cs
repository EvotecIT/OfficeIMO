using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal static class OpenXmlWorkbookElementOrder {
        internal static void InsertInOrder(Workbook workbook, OpenXmlElement element) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (element == null) throw new ArgumentNullException(nameof(element));

            if (element is WorkbookProperties) {
                OpenXmlElement? before = workbook.GetFirstChild<WorkbookProtection>();
                before ??= workbook.GetFirstChild<BookViews>();
                before ??= workbook.GetFirstChild<Sheets>();
                if (before != null) {
                    workbook.InsertBefore(element, before);
                } else {
                    workbook.Append(element);
                }
                return;
            }

            workbook.Append(element);
        }
    }
}
