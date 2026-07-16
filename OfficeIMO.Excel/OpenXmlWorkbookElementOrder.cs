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

            if (element is WorkbookProtection) {
                OpenXmlElement? before = workbook.GetFirstChild<BookViews>();
                before ??= workbook.GetFirstChild<Sheets>();
                if (before != null) {
                    workbook.InsertBefore(element, before);
                } else if (workbook.GetFirstChild<WorkbookProperties>() is WorkbookProperties properties) {
                    workbook.InsertAfter(element, properties);
                } else {
                    workbook.Append(element);
                }
                return;
            }

            if (element is CalculationProperties) {
                OpenXmlElement? before = workbook.ChildElements.FirstOrDefault(child =>
                    string.Equals(child.LocalName, "oleSize", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "customWorkbookViews", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "pivotCaches", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "smartTagPr", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "smartTagTypes", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "webPublishing", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "fileRecoveryPr", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "webPublishObjects", StringComparison.Ordinal)
                    || string.Equals(child.LocalName, "extLst", StringComparison.Ordinal));
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
