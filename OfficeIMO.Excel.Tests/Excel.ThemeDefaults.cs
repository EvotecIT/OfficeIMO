using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void TypedLineChartsPreserveExplicitMarkersAndOtherwiseUseTheTheme(bool explicitMarkers) {
            double[] values = { 1D, 2D, 3D };
            OfficeSeries series = explicitMarkers
                ? new OfficeSeries("Sales", values, null, null, null, true)
                : new OfficeSeries("Sales", values);
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Sales").AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] { series }), new ExcelChartOptions { Row = 1, Column = 1 });
            using var stream = new MemoryStream(document.ToBytes());
            using SpreadsheetDocument saved = SpreadsheetDocument.Open(stream, false);
            C.Marker marker = saved.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null)
                .DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Marker>()!;

            Assert.Equal(!explicitMarkers, marker.Symbol!.Val!.Value == C.MarkerStyleValues.None);
        }

        [Fact]
        public void TabularSavePreservesAnAuthoredWorkbookTheme() {
            using ExcelDocument document = ExcelDocument.Create();
            document.ResetWorkbookTheme("Regional reporting");
            var data = new DataTable("Data");
            data.Columns.Add("Region");
            data.Rows.Add("Europe");
            var set = new DataSet();
            set.Tables.Add(data);
            document.InsertDataSet(set, createTables: false, includeAutoFilter: false);
            using var stream = new MemoryStream(document.ToBytes());
            using ExcelDocument reopened = ExcelDocument.Load(stream);

            Assert.Equal("Regional reporting", reopened.GetWorkbookTheme().Name);
            Assert.True(reopened.Sheets.Single().TryGetCellText(2, 1, out string? region));
            Assert.Equal("Europe", region);
        }

        [Fact]
        public void NewChartsWriteSharedDefaultsAndRetainExplicitTextOverrides() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelChart chart = document.AddSheet("Data").AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "East", "West" }, new[] { new ExcelChartSeries("Sales", new[] { 30D, 42D }) }), new ExcelChartOptions { Row = 1, Column = 1, Title = "Sales by region" });
            using (var stream = new MemoryStream(document.ToBytes()))
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false)) {
                C.ChartSpace space = package.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single().ChartSpace!;
                A.DefaultRunProperties font = Assert.Single(space.GetFirstChild<C.TextProperties>()!.Descendants<A.DefaultRunProperties>());
                Assert.Equal(OfficeTheme.Default.ChartStyle.AxisFont.Family, font.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
                Assert.Equal((int)(OfficeTheme.Default.ChartStyle.AxisFontSize.ToPoints() * 100D), font.FontSize!.Value);
                Assert.Equal(OfficeTheme.Default.ChartStyle.BarGapPercent, space.Descendants<C.GapWidth>().Single().Val!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(package));
            }
            chart.SetTitleTextStyle(fontSize: OfficeLength.Points(17D), fontName: "Georgia");
            chart.SetCategoryAxisLabelTextStyle(fontSize: OfficeLength.Points(10D), fontName: "Courier New");
            using var saved = new MemoryStream(document.ToBytes());
            using SpreadsheetDocument reopened = SpreadsheetDocument.Open(saved, false);
            C.ChartSpace result = reopened.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single().ChartSpace!;
            Assert.Contains(result.Descendants<C.Title>().Single().Descendants<A.LatinFont>(), font => font.Typeface!.Value == "Georgia");
            Assert.Contains(result.Descendants<C.CategoryAxis>().Single().Descendants<A.LatinFont>(), font => font.Typeface!.Value == "Courier New");
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void NewWorkbooksUseSharedThemeInEditableAndTabularExports(bool tabular) {
            using ExcelDocument document = ExcelDocument.Create();
            if (tabular) {
                var data = new DataTable("Data");
                data.Columns.Add("Region");
                data.Rows.Add("Europe");
                var set = new DataSet();
                set.Tables.Add(data);
                document.InsertDataSet(set, createTables: false, includeAutoFilter: false);
            } else {
                document.AddSheet("Data").CellValue(1, 1, "Europe");
            }

            using var stream = new MemoryStream(document.ToBytes());
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            var workbook = package.WorkbookPart!;
            Font font = workbook.WorkbookStylesPart!.Stylesheet.Fonts!.Elements<Font>().First();
            Assert.Equal(OfficeTheme.Default.Fonts.Body.Family, font.FontName!.Val!.Value);
            Assert.Equal(OfficeTheme.Default.TypeScale.Body.ToPoints(), font.FontSize!.Val!.Value);
            using var themeStream = workbook.ThemePart!.GetStream();
            XElement theme = XDocument.Load(themeStream).Root!;
            XNamespace ns = theme.Name.Namespace;
            XElement elements = theme.Element(ns + "themeElements")!;
            XElement colors = elements.Element(ns + "clrScheme")!;
            Assert.Equal(OfficeTheme.Default.Palette.Accent1.ToRgbHex(),
                colors.Element(ns + "accent1")!.Element(ns + "srgbClr")!.Attribute("val")!.Value);
            Assert.Equal(OfficeTheme.Default.Fonts.Heading.Family,
                elements.Element(ns + "fontScheme")!.Element(ns + "majorFont")!.Element(ns + "latin")!.Attribute("typeface")!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
    }
}
