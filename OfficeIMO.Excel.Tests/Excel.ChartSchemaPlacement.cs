using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // Charts saved by OfficeIMO, then edited through the SDK to hold markup other producers write, then loaded again.
        private static byte[] CreateChartWorkbook(OfficeChartType type, Action<C.ChartSpace> edit) {
            using var created = new MemoryStream();
            using (var document = ExcelDocument.Create(created)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.AddChart(type, ExcelChartData.CreateDefault(), new ExcelChartOptions { Row = 2, Column = 2 });
                document.Save();
            }
            using var edited = new MemoryStream();
            created.Position = 0;
            created.CopyTo(edited);
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(edited, true)) {
                ChartPart chartPart = GetWorksheetPartWithCharts(spreadsheet).DrawingsPart!.ChartParts.Single();
                edit(chartPart.ChartSpace);
                chartPart.ChartSpace.Save();
            }
            return edited.ToArray();
        }

        private static C.ChartSpace ReadChartSpace(byte[] package) {
            using var stream = new MemoryStream(package, writable: false);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
            ChartPart chartPart = GetWorksheetPartWithCharts(spreadsheet).DrawingsPart!.ChartParts.Single();
            return (C.ChartSpace)chartPart.ChartSpace.CloneNode(true);
        }

        private static byte[] EditChart(byte[] package, Action<ExcelChart> edit) {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(package))) {
                edit(document.Sheets.First().Charts.Single());
                document.Save(output);
            }
            return output.ToArray();
        }

        [Fact]
        public void Test_Chart_DataLabelNumberFormatReplacesDeletedLabels() {
            byte[] package = CreateChartWorkbook(OfficeChartType.ColumnClustered, space => {
                C.BarChart bars = space.Descendants<C.BarChart>().Single();
                bars.RemoveAllChildren<C.DataLabels>();
                bars.InsertBefore(new C.DataLabels(new C.Delete { Val = true }), bars.GetFirstChild<C.GapWidth>());
            });

            byte[] saved = EditChart(package, chart => chart.SetDataLabels(numberFormat: "0.0"));

            C.DataLabels labels = ReadChartSpace(saved).Descendants<C.BarChart>().Single().GetFirstChild<C.DataLabels>()!;
            Assert.Null(labels.GetFirstChild<C.Delete>());
            Assert.Equal("0.0", labels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
        }

        [Fact]
        public void Test_Chart_UpdateDataPutsNewLineSeriesBeforeChartSmooth() {
            byte[] package = CreateChartWorkbook(OfficeChartType.Line, space => {
                C.LineChart line = space.Descendants<C.LineChart>().Single();
                foreach (C.LineChartSeries extra in line.Elements<C.LineChartSeries>().Skip(1).ToList()) extra.Remove();
                line.RemoveAllChildren<C.DataLabels>();
                line.RemoveAllChildren<C.ShowMarker>();
                line.RemoveAllChildren<C.Smooth>();
                line.InsertBefore(new C.Smooth { Val = false }, line.GetFirstChild<C.AxisId>());
            });

            byte[] saved = EditChart(package, chart => chart.UpdateData(ExcelChartData.CreateDefault()));

            string[] children = ReadChartSpace(saved).Descendants<C.LineChart>().Single().ChildElements
                .Select(child => child.LocalName).ToArray();
            int lastSeries = Array.LastIndexOf(children, "ser");
            Assert.Equal(3, children.Count(name => name == "ser"));
            Assert.True(lastSeries < Array.IndexOf(children, "smooth"), string.Join(",", children));
        }

        [Fact]
        public void Test_Chart_UpdateDataKeepsLineSeriesDataAfterPictureOptions() {
            byte[] package = CreateChartWorkbook(OfficeChartType.Line, space => {
                foreach (C.LineChartSeries series in space.Descendants<C.LineChartSeries>()) {
                    series.InsertBefore(new C.PictureOptions(), series.GetFirstChild<C.CategoryAxisData>());
                }
            });

            byte[] saved = EditChart(package, chart => chart.UpdateData(ExcelChartData.CreateDefault()));

            foreach (C.LineChartSeries series in ReadChartSpace(saved).Descendants<C.LineChartSeries>()) {
                string[] children = series.ChildElements.Select(child => child.LocalName).ToArray();
                int pictureOptions = Array.IndexOf(children, "pictureOptions");
                Assert.True(pictureOptions >= 0, string.Join(",", children));
                Assert.True(pictureOptions < Array.IndexOf(children, "cat"), string.Join(",", children));
            }
        }
    }
}
