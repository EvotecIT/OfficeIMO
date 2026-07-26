using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        private static bool HasUnsupportedBubbleSourceVisibility(
            ChartPart chartPart, C.Chart chart) {
            C.PlotVisibleOnly? plotVisibleOnly =
                chart.GetFirstChild<C.PlotVisibleOnly>();
            if (plotVisibleOnly == null ||
                plotVisibleOnly.Val?.Value == false) {
                return false;
            }

            EmbeddedPackagePart? embedded = chartPart
                .GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();
            if (embedded == null) return true;
            try {
                using Stream stream = embedded.GetStream(
                    FileMode.Open, FileAccess.Read);
                using SpreadsheetDocument workbook =
                    SpreadsheetDocument.Open(stream, false);
                return workbook.WorkbookPart?.WorksheetParts.Any(part =>
                    part.Worksheet?.Descendants<S.Row>().Any(row =>
                        row.Hidden?.Value == true) == true ||
                    part.Worksheet?.Descendants<S.Column>().Any(column =>
                        column.Hidden?.Value == true) == true) != false;
            } catch {
                return true;
            }
        }
    }
}
