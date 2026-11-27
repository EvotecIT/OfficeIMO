using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        internal static byte[] BuildBubbleChartWorkbook(OfficeChartData data) {
            ArgumentNullException.ThrowIfNull(data);

            using MemoryStream stream = new();
            using (SpreadsheetDocument document =
                   SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new S.Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new S.SheetData();

                int seriesCount = data.Series.Count;
                int totalColumns = seriesCount * 3;
                int maxPoints = data.Series.Max(series => series.Values.Count);
                int totalRows = maxPoints + 1;
                worksheetPart.Worksheet = new S.Worksheet(
                    new S.SheetDimension {
                        Reference = $"A1:{ColumnLetter(totalColumns)}{totalRows}"
                    },
                    sheetData);

                SharedStringTablePart sharedStringsPart =
                    workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringsPart.SharedStringTable = new S.SharedStringTable();
                var stringIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
                int GetStringIndex(string value) {
                    if (!stringIndexes.TryGetValue(value, out int index)) {
                        index = stringIndexes.Count;
                        stringIndexes[value] = index;
                        sharedStringsPart.SharedStringTable.AppendChild(
                            new S.SharedStringItem(new S.Text(value)));
                    }
                    return index;
                }

                var header = new S.Row {
                    RowIndex = 1U,
                    Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" }
                };
                for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                    OfficeSeries series = data.Series[seriesIndex];
                    int xColumn = seriesIndex * 3 + 1;
                    header.Append(CreateSharedStringCell($"{ColumnLetter(xColumn)}1",
                        GetStringIndex($"{series.Name} X")));
                    header.Append(CreateSharedStringCell($"{ColumnLetter(xColumn + 1)}1",
                        GetStringIndex(series.Name)));
                    header.Append(CreateSharedStringCell($"{ColumnLetter(xColumn + 2)}1",
                        GetStringIndex($"{series.Name} Size")));
                }
                sheetData.Append(header);

                IReadOnlyList<double> sharedX = data.Series.Any(series => series.XValues == null)
                    ? ParseScatterCategories(data.Categories)
                    : Array.Empty<double>();
                var rows = new S.Row[maxPoints];
                for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                    OfficeSeries series = data.Series[seriesIndex];
                    IReadOnlyList<double> xValues = series.XValues ?? sharedX;
                    int xColumn = seriesIndex * 3 + 1;
                    for (int pointIndex = 0;
                         pointIndex < series.Values.Count;
                         pointIndex++) {
                        uint rowIndex = (uint)(pointIndex + 2);
                        S.Row row = rows[pointIndex] ??= new S.Row {
                            RowIndex = rowIndex,
                            Spans = new ListValue<StringValue> {
                                InnerText = $"1:{totalColumns}"
                            }
                        };
                        row.Append(CreateNumberCell(
                            $"{ColumnLetter(xColumn)}{rowIndex}",
                            xValues[pointIndex]));
                        row.Append(CreateNumberCell(
                            $"{ColumnLetter(xColumn + 1)}{rowIndex}",
                            series.Values[pointIndex]));
                        row.Append(CreateNumberCell(
                            $"{ColumnLetter(xColumn + 2)}{rowIndex}",
                            series.BubbleSizes![pointIndex]));
                    }
                }
                foreach (S.Row row in rows) {
                    sheetData.Append(row);
                }

                S.Sheets sheets = workbookPart.Workbook.AppendChild(new S.Sheets());
                sheets.Append(new S.Sheet {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "Sheet1"
                });
                sharedStringsPart.SharedStringTable.Save();
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
            return stream.ToArray();
        }

        private static OfficeXmlNode CreateSharedBubbleChart(OfficeXmlNode context, OfficeChartData data,
            uint xAxisId, uint yAxisId) {
            OfficeXmlNode chart = context.CreateElement(NsC, "bubbleChart");
            chart.AppendVal(NsC, "varyColors", Text(false));
            for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
                chart.Append(CreateSharedBubbleSeries(chart, data, seriesIndex));
            }
            chart.Append(CreateDefaultDataLabels(chart));
            chart.AppendVal(NsC, "bubble3D", Text(false));
            chart.AppendVal(NsC, "bubbleScale", Text(100U));
            chart.AppendVal(NsC, "showNegBubbles", Text(false));
            chart.AppendVal(NsC, "sizeRepresents", "area");
            chart.AppendVal(NsC, "axId", Text(xAxisId));
            chart.AppendVal(NsC, "axId", Text(yAxisId));
            return chart;
        }

        private static OfficeXmlNode CreateSharedBubbleSeries(OfficeXmlNode context, OfficeChartData data,
            int seriesIndex) {
            OfficeSeries series = data.Series[seriesIndex];
            IReadOnlyList<double> xValues = series.XValues ?? ParseScatterCategories(data.Categories);
            int firstColumn = seriesIndex * 3 + 1;
            int lastRow = series.Values.Count + 1;
            string xColumn = ColumnLetter(firstColumn);
            string yColumn = ColumnLetter(firstColumn + 1);
            string sizeColumn = ColumnLetter(firstColumn + 2);
            OfficeXmlNode seriesElement = context.CreateElement(NsC, "ser");
            seriesElement.AppendVal(NsC, "idx", Text((uint)seriesIndex));
            seriesElement.AppendVal(NsC, "order", Text((uint)seriesIndex));
            AppendStringReference(seriesElement.AppendNew(NsC, "tx"),
                $"Sheet1!${yColumn}$1", new[] { series.Name });
            AppendNumberReference(seriesElement.AppendNew(NsC, "xVal"),
                $"Sheet1!${xColumn}$2:${xColumn}${lastRow}", xValues);
            AppendNumberReference(seriesElement.AppendNew(NsC, "yVal"),
                $"Sheet1!${yColumn}$2:${yColumn}${lastRow}", series.Values);
            AppendNumberReference(seriesElement.AppendNew(NsC, "bubbleSize"),
                $"Sheet1!${sizeColumn}$2:${sizeColumn}${lastRow}", series.BubbleSizes!);
            seriesElement.AppendVal(NsC, "bubble3D", Text(false));
            return seriesElement;
        }
    }
}
