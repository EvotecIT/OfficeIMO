using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;
using OfficeReferenceSequence = DocumentFormat.OpenXml.Office.Excel.ReferenceSequence;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetSparklineResolver {
        internal static IReadOnlyList<ExcelWorksheetSparklineInfo> FindSparklines(WorksheetPart worksheetPart) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            if (worksheetPart.Worksheet == null) {
                return Array.Empty<ExcelWorksheetSparklineInfo>();
            }

            var sparklines = new List<ExcelWorksheetSparklineInfo>();
            int groupIndex = 0;
            foreach (SparklineGroup group in worksheetPart.Worksheet.Descendants<SparklineGroup>()) {
                string kind = group.Type?.InnerText ?? string.Empty;
                foreach (Sparkline sparkline in group.Descendants<Sparkline>()) {
                    string location = sparkline.GetFirstChild<OfficeReferenceSequence>()?.Text ?? string.Empty;
                    string formula = sparkline.GetFirstChild<OfficeFormula>()?.Text ?? string.Empty;
                    foreach (string cellReference in ExpandLocation(location)) {
                        sparklines.Add(new ExcelWorksheetSparklineInfo(
                            groupIndex,
                            cellReference,
                            formula,
                            kind,
                            group.Markers?.Value == true,
                            group.High?.Value == true,
                            group.Low?.Value == true,
                            group.First?.Value == true,
                            group.Last?.Value == true,
                            group.Negative?.Value == true,
                            group.DisplayXAxis?.Value == true,
                            NormalizeColor(group.SeriesColor?.Rgb?.Value),
                            NormalizeColor(group.AxisColor?.Rgb?.Value),
                            NormalizeColor(group.NegativeColor?.Rgb?.Value),
                            NormalizeColor(group.MarkersColor?.Rgb?.Value),
                            NormalizeColor(group.HighMarkerColor?.Rgb?.Value),
                            NormalizeColor(group.LowMarkerColor?.Rgb?.Value),
                            NormalizeColor(group.FirstMarkerColor?.Rgb?.Value),
                            NormalizeColor(group.LastMarkerColor?.Rgb?.Value)));
                    }
                }

                groupIndex++;
            }

            return sparklines;
        }

        private static IEnumerable<string> ExpandLocation(string location) {
            string normalized = NormalizeReference(location);
            if (A1.TryParseRange(normalized, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                for (int row = firstRow; row <= lastRow; row++) {
                    for (int column = firstColumn; column <= lastColumn; column++) {
                        yield return A1.CellReference(row, column);
                    }
                }

                yield break;
            }

            (int singleRow, int singleColumn) = A1.ParseCellRef(normalized);
            if (singleRow > 0 && singleColumn > 0) {
                yield return A1.CellReference(singleRow, singleColumn);
            }
        }

        private static string NormalizeReference(string reference) {
            string value = (reference ?? string.Empty).Trim();
            int separator = value.LastIndexOf('!');
            if (separator >= 0 && separator + 1 < value.Length) {
                value = value.Substring(separator + 1);
            }

            return value.Replace("$", string.Empty);
        }

        private static string? NormalizeColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string color = value!.Trim().TrimStart('#');
            if (color.Length == 6) {
                return "FF" + color.ToUpperInvariant();
            }

            return color.Length == 8 ? color.ToUpperInvariant() : null;
        }
    }

    internal sealed class ExcelWorksheetSparklineInfo {
        internal ExcelWorksheetSparklineInfo(
            int groupIndex,
            string cellReference,
            string formula,
            string kind,
            bool displayMarkers,
            bool displayHigh,
            bool displayLow,
            bool displayFirst,
            bool displayLast,
            bool displayNegative,
            bool displayAxis,
            string? seriesColorArgb,
            string? axisColorArgb,
            string? negativeColorArgb,
            string? markersColorArgb,
            string? highColorArgb,
            string? lowColorArgb,
            string? firstColorArgb,
            string? lastColorArgb) {
            GroupIndex = groupIndex;
            CellReference = cellReference ?? string.Empty;
            Formula = formula ?? string.Empty;
            Kind = kind ?? string.Empty;
            DisplayMarkers = displayMarkers;
            DisplayHigh = displayHigh;
            DisplayLow = displayLow;
            DisplayFirst = displayFirst;
            DisplayLast = displayLast;
            DisplayNegative = displayNegative;
            DisplayAxis = displayAxis;
            SeriesColorArgb = seriesColorArgb;
            AxisColorArgb = axisColorArgb;
            NegativeColorArgb = negativeColorArgb;
            MarkersColorArgb = markersColorArgb;
            HighColorArgb = highColorArgb;
            LowColorArgb = lowColorArgb;
            FirstColorArgb = firstColorArgb;
            LastColorArgb = lastColorArgb;
        }

        internal int GroupIndex { get; }

        internal string CellReference { get; }

        internal string Formula { get; }

        internal string Kind { get; }

        internal bool DisplayMarkers { get; }

        internal bool DisplayHigh { get; }

        internal bool DisplayLow { get; }

        internal bool DisplayFirst { get; }

        internal bool DisplayLast { get; }

        internal bool DisplayNegative { get; }

        internal bool DisplayAxis { get; }

        internal string? SeriesColorArgb { get; }

        internal string? AxisColorArgb { get; }

        internal string? NegativeColorArgb { get; }

        internal string? MarkersColorArgb { get; }

        internal string? HighColorArgb { get; }

        internal string? LowColorArgb { get; }

        internal string? FirstColorArgb { get; }

        internal string? LastColorArgb { get; }
    }
}
