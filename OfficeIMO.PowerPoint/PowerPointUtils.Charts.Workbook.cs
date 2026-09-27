using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        internal static byte[] BuildChartWorkbook(PowerPointChartData data) =>
            OfficeOpenXmlChartWriter.BuildWorkbook(
                new OfficeChartData(data.Categories, data.Series.Select(series =>
                    new OfficeChartSeries(series.Name, series.Values))),
                OfficeChartKind.ColumnClustered);

        internal static byte[] BuildChartWorkbook(PowerPointScatterChartData data) =>
            OfficeOpenXmlChartWriter.BuildWorkbook(
                new OfficeChartData(data.Series[0].XValues.Select(value => value.ToString(CultureInfo.InvariantCulture)),
                    data.Series.Select(series => new OfficeChartSeries(series.Name, series.YValues, series.XValues))),
                OfficeChartKind.Scatter);

        private static string ColumnLetter(int column) => OfficeOpenXmlChartWriter.ColumnLetter(column);
    }
}