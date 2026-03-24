using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const uint MaxExcelRowIndex = 1_048_575U;
        private const uint MaxExcelColumnIndex = 16_383U;

        internal void CleanupPrintArtifacts() {
            var worksheet = WorksheetRoot;

            CleanupRowBreaks(worksheet.GetFirstChild<RowBreaks>());
            CleanupColumnBreaks(worksheet.GetFirstChild<ColumnBreaks>());

            var printOptions = worksheet.GetFirstChild<PrintOptions>();
            if (printOptions != null && !printOptions.HasAttributes) {
                printOptions.Remove();
            }

            var pageMargins = worksheet.GetFirstChild<PageMargins>();
            if (pageMargins != null) {
                pageMargins.Left = NormalizeMargin(pageMargins.Left, 0.7D);
                pageMargins.Right = NormalizeMargin(pageMargins.Right, 0.7D);
                pageMargins.Top = NormalizeMargin(pageMargins.Top, 0.75D);
                pageMargins.Bottom = NormalizeMargin(pageMargins.Bottom, 0.75D);
                pageMargins.Header = NormalizeMargin(pageMargins.Header, 0.3D);
                pageMargins.Footer = NormalizeMargin(pageMargins.Footer, 0.3D);
            }

            var pageSetup = worksheet.GetFirstChild<PageSetup>();
            if (pageSetup?.Scale?.Value is uint scale && (scale < 10U || scale > 400U)) {
                pageSetup.Scale = 100U;
            }
        }

        private static void CleanupRowBreaks(RowBreaks? rowBreaks) {
            if (rowBreaks == null) {
                return;
            }

            NormalizeBreakContainer(
                rowBreaks,
                maxIdInclusive: MaxExcelRowIndex,
                defaultMax: MaxExcelColumnIndex,
                setCounts: count => {
                    rowBreaks.Count = count;
                    rowBreaks.ManualBreakCount = count;
                });
        }

        private static void CleanupColumnBreaks(ColumnBreaks? columnBreaks) {
            if (columnBreaks == null) {
                return;
            }

            NormalizeBreakContainer(
                columnBreaks,
                maxIdInclusive: MaxExcelColumnIndex,
                defaultMax: MaxExcelRowIndex,
                setCounts: count => {
                    columnBreaks.Count = count;
                    columnBreaks.ManualBreakCount = count;
                });
        }

        private static void NormalizeBreakContainer(OpenXmlCompositeElement container, uint maxIdInclusive, uint defaultMax, Action<uint> setCounts) {
            var seenIds = new HashSet<uint>();
            foreach (var pageBreak in container.Elements<Break>().ToList()) {
                if (pageBreak.Id?.Value is not uint id || id > maxIdInclusive || !seenIds.Add(id)) {
                    pageBreak.Remove();
                    continue;
                }

                uint min = pageBreak.Min?.Value ?? 0U;
                uint max = pageBreak.Max?.Value ?? defaultMax;
                if (min > defaultMax || max > defaultMax || min > max) {
                    min = 0U;
                    max = defaultMax;
                }

                pageBreak.Min = min;
                pageBreak.Max = max;
                pageBreak.ManualPageBreak = true;
            }

            uint remaining = (uint)container.Elements<Break>().Count();
            if (remaining == 0) {
                container.Remove();
                return;
            }

            setCounts(remaining);
        }

        private static DoubleValue NormalizeMargin(DoubleValue? margin, double fallback) {
            if (margin?.Value == null || double.IsNaN(margin.Value) || double.IsInfinity(margin.Value) || margin.Value < 0D) {
                return fallback;
            }

            return margin;
        }
    }
}
