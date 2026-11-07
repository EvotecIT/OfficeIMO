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
        [Fact]
        public void NewWorkbookThemesPreserveDistinctFontsAndPalettesWithTheSameName() {
            OfficeTheme first = OfficeTheme.Modern
                .WithPalette(OfficeTheme.Modern.Palette.With(accent1: OfficeColor.ParseHex("123456")))
                .WithFonts(OfficeTheme.Modern.Fonts.With(body: OfficeTheme.Modern.Fonts.Body.With(family: "Arial")));
            OfficeTheme second = first
                .WithPalette(first.Palette.With(accent1: OfficeColor.ParseHex("ABCDEF")))
                .WithFonts(first.Fonts.With(body: first.Fonts.Body.With(family: "Georgia")));
            Assert.Equal(first.Name, second.Name);
            foreach (var item in new[] { (Theme: first, Font: "Arial", Accent: "123456"),
                (Theme: second, Font: "Georgia", Accent: "ABCDEF"), (Theme: first, Font: "Arial", Accent: "123456") }) {
                using ExcelDocument document = ExcelDocument.Create(new ExcelCreateOptions { Theme = item.Theme });
                document.AddSheet("Data").WriteData(new[] { 1 }, ("Value", static value => value));
                using var stream = new MemoryStream(document.ToBytes());
                using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
                using var themeStream = package.WorkbookPart!.ThemePart!.GetStream();
                XElement theme = XDocument.Load(themeStream).Root!;
                XNamespace drawing = theme.Name.Namespace;
                XElement elements = theme.Element(drawing + "themeElements")!;
                Assert.Equal(first.Name, (string?)theme.Attribute("name"));
                Assert.Equal(item.Accent, (string?)elements.Element(drawing + "clrScheme")!
                    .Element(drawing + "accent1")!.Element(drawing + "srgbClr")!.Attribute("val"));
                Assert.Equal(item.Font, (string?)elements.Element(drawing + "fontScheme")!
                    .Element(drawing + "minorFont")!.Element(drawing + "latin")!.Attribute("typeface"));
            }
        }

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
            chart.SetTitleTextStyle(fontSize: OfficeLength.Points(17D), fontFamily: "Georgia");
            chart.SetCategoryAxisLabelTextStyle(fontSize: OfficeLength.Points(10D), fontFamily: "Courier New");
            using var saved = new MemoryStream(document.ToBytes());
            using SpreadsheetDocument reopened = SpreadsheetDocument.Open(saved, false);
            C.ChartSpace result = reopened.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single().ChartSpace!;
            Assert.Contains(result.Descendants<C.Title>().Single().Descendants<A.LatinFont>(), font => font.Typeface!.Value == "Georgia");
            Assert.Contains(result.Descendants<C.CategoryAxis>().Single().Descendants<A.LatinFont>(), font => font.Typeface!.Value == "Courier New");
        }

        // Font zero is the Normal style font, whose digit width sizes every column, so it must be
        // the creation theme's body font whichever package writer saves the content.
        [Theory]
        [InlineData("cells", false, false, false)]
        [InlineData("dataSet", false, false, false)]
        [InlineData("dataWithChart", false, false, false)]
        [InlineData("cells", true, false, false)]
        [InlineData("dataSet", true, false, false)]
        [InlineData("dataWithChart", true, false, false)]
        [InlineData("cells", true, true, false)]
        [InlineData("dataSet", true, true, false)]
        [InlineData("dataWithChart", true, true, false)]
        [InlineData("dataSet", true, true, true)]
        public void NewWorkbooksPreserveTheCreationFontAcrossSaveRoutes(string content, bool themed, bool applyTheme, bool disableFastWriter) {
            OfficeTheme? creationTheme = themed ? OfficeTheme.Minimal : null;
            OfficeTheme expected = creationTheme ?? OfficeTheme.Default;
            OfficeTheme workbookTheme = applyTheme
                ? OfficeTheme.Classic.WithTypeScale(OfficeTheme.Classic.TypeScale.With(body: OfficeLength.Points(19D)))
                : expected;
            using ExcelDocument document = ExcelDocument.Create(new ExcelCreateOptions { Theme = creationTheme });
            if (applyTheme) document.ApplyTheme(workbookTheme);
            if (content == "dataSet") {
                var data = new DataTable("Data");
                data.Columns.Add("Region");
                data.Rows.Add("Europe");
                var set = new DataSet();
                set.Tables.Add(data);
                document.InsertDataSet(set, createTables: false, includeAutoFilter: false);
            } else if (content == "dataWithChart") {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.WriteData(new[] { 30D, 42D }, ("Sales", static _ => _));
                sheet.AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "East", "West" }, new[] { new ExcelChartSeries("Sales", new[] { 30D, 42D }) }), new ExcelChartOptions { Row = 5, Column = 3 });
            } else {
                document.AddSheet("Data").Cell(1, 1).SetValue("Europe");
            }

            using var stream = new MemoryStream(document.ToBytes(new ExcelSaveOptions { DisableFastPackageWriter = disableFastWriter }));
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            var workbook = package.WorkbookPart!;
            Font font = workbook.WorkbookStylesPart!.Stylesheet.Fonts!.Elements<Font>().First();
            Assert.Equal(expected.Fonts.Body.Family, font.FontName!.Val!.Value);
            Assert.Equal(expected.TypeScale.Body.ToPoints(), font.FontSize!.Val!.Value);
            Assert.Empty(font.Elements<FontScheme>());
            using var themeStream = workbook.ThemePart!.GetStream();
            XElement theme = XDocument.Load(themeStream).Root!;
            XNamespace ns = theme.Name.Namespace;
            XElement elements = theme.Element(ns + "themeElements")!;
            XElement colors = elements.Element(ns + "clrScheme")!;
            Assert.Equal(workbookTheme.Palette.Accent1.ToRgbHex(),
                colors.Element(ns + "accent1")!.Element(ns + "srgbClr")!.Attribute("val")!.Value);
            XElement fonts = elements.Element(ns + "fontScheme")!;
            Assert.Equal(workbookTheme.Fonts.Heading.Family,
                fonts.Element(ns + "majorFont")!.Element(ns + "latin")!.Attribute("typeface")!.Value);
            Assert.Equal(workbookTheme.Fonts.Body.Family,
                fonts.Element(ns + "minorFont")!.Element(ns + "latin")!.Attribute("typeface")!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void StreamedWorkbooksUseTheWriterThemeForTheNormalFont(bool declaredStyles) {
            ExcelWriterOptions? options = declaredStyles
                ? new ExcelWriterOptions {
                    UseSharedStrings = false,
                    Styles = new Dictionary<string, ExcelStyleDefinition> { ["Total"] = new() { Bold = true } },
                    DefaultRowStyle = "Total"
                }
                : null;
            using var output = new MemoryStream();
            using (var writer = ExcelDocument.CreateWriter(output, options, OfficeTheme.Minimal, leaveOpen: true)) {
                writer.WriteTable("Data", new[] { 1, 2 }, new[] { "Value" }, (row, value) => row.Write(value));
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Font font = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet.Fonts!.Elements<Font>().First();
            Assert.Equal(OfficeTheme.Minimal.Fonts.Body.Family, font.FontName!.Val!.Value);
            Assert.Equal(OfficeTheme.Minimal.TypeScale.Body.ToPoints(), font.FontSize!.Val!.Value);
            Assert.Empty(font.Elements<FontScheme>());
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
    }
}
