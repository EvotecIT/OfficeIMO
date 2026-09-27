using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        private static bool HasUnsupportedBubbleSourceVisibility(ChartPart part, C.Chart chart) {
            var visibleOnly = chart.GetFirstChild<C.PlotVisibleOnly>();
            if (visibleOnly == null || visibleOnly.Val?.Value == false) return false;
            try {
                var embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(part);
                if (embedded == null) return true;
                using var stream = embedded.GetStream(FileMode.Open, FileAccess.Read);
                byte[] bytes = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(stream);
                using var workbookStream = new MemoryStream(bytes, writable: false);
                using var workbook = SpreadsheetDocument.Open(workbookStream, false, OfficeOpenXmlChartWorkbookSecurity.CreateOpenSettings());
                return workbook.WorkbookPart?.WorksheetParts.Any(sheet =>
                    sheet.Worksheet?.Descendants<S.Row>().Any(row => row.Hidden?.Value == true) == true ||
                    sheet.Worksheet?.Descendants<S.Column>().Any(column => column.Hidden?.Value == true) == true) != false;
            } catch {
                return true;
            }
        }
    }
}
