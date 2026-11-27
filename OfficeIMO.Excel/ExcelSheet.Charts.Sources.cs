using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void EnsureChartAxisIds() => ExcelChartAxisIdGenerator.EnsureInitialized(ExcelPartRef.WorksheetForMetadata(this));

        private static void PopulateChart(ExcelPartRef chartPart, OfficeChartType type, ExcelChartDataRange range,
            ExcelChartData? data, string? title, IReadOnlyList<string>? nativeSeriesNames) =>
            ExcelChartUtils.PopulateChart(chartPart, type, range, data, title, nativeSeriesNames);

        private static void PopulateChartFromSeriesRanges(ExcelPartRef chartPart, OfficeChartType type, string defaultSheetName,
            IReadOnlyList<ExcelChartSeriesRange> ranges, string? title) =>
            ExcelChartUtils.PopulateChartFromSeriesRanges(chartPart, type, defaultSheetName, ranges, title);

        private static void ApplyChartStyle(ExcelPartRef chartPart, ExcelChartStylePreset preset) =>
            ExcelChartUtils.ApplyChartStyle(chartPart, preset);

        private static void ApplyRangeSeriesFormulas(ExcelChartRangeSeries rangeSeries, ExcelPartRef chartPart) =>
            rangeSeries.ApplyFormulas(chartPart);

        /// <summary>The parts of a table definition a chart is sourced from.</summary>
        private sealed class ChartSourceTable {
            internal string? Reference { get; set; }
            internal uint? HeaderRowCount { get; set; }
            internal uint? TotalsRowCount { get; set; }
            internal IReadOnlyList<string>? ColumnNames { get; set; }
        }

        /// <summary>Reads the worksheet's table definition with this name (or display name), in relationship order.</summary>
        private ChartSourceTable? FindChartSourceTable(string tableName) {
            foreach (ExcelPartRef part in ExcelPartRef.WorksheetForMetadata(this).PartsOfKind(ExcelPartKind.Table)) {
                OfficeXmlNode table = part.Root;
                if (table.IsNone) continue;
                string? name = table.Attribute("name") ?? table.Attribute("displayName");
                if (!string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase)) continue;
                OfficeXmlNode columns = table.Child(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.SpreadsheetML, "tableColumns");
                List<string>? names = null;
                if (columns.Exists) {
                    names = new List<string>();
                    foreach (OfficeXmlNode column in columns.Children(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.SpreadsheetML, "tableColumn")) {
                        names.Add(column.Attribute("name") ?? string.Empty);
                    }
                }
                return new ChartSourceTable {
                    Reference = table.Attribute("ref"),
                    HeaderRowCount = table.UInt32Attribute("headerRowCount"),
                    TotalsRowCount = table.UInt32Attribute("totalsRowCount"),
                    ColumnNames = names?.ToArray()
                };
            }
            return null;
        }
    }
}
