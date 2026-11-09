using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Tests;

public sealed class ExcelFormattingAuthoringTests {
    [Fact]
    public void HeaderSettersClearFlagsAndPreserveOtherFormatting() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Formatting");
        sheet.Cell(1, 1).SetValue("Status");
        sheet.Cell(2, 1).SetValue("Ready");
        sheet.Cell(3, 1).SetValue("Pending");
        sheet.ColumnStyleByHeader("Status").SetBold().SetItalic().SetUnderline(ExcelUnderlineStyle.DoubleAccounting)
            .SetStrikethrough().SetWrapText().SetFontFamily("Aptos").SetFontSize(OfficeLength.Points(14))
            .SetColor(OfficeColor.Accent1).SetFillColor(OfficeColor.Accent2).SetNumberFormat("@").SetHorizontalAlignment(ExcelHorizontalAlignment.Right)
            .SetVerticalAlignment(ExcelVerticalAlignment.Center).SetVerticalTextAlignment(ExcelVerticalTextAlignment.Superscript)
            .SetBold(false).SetItalic(false).SetUnderline(false).SetStrikethrough(false).SetWrapText(false);
        sheet.ColumnStyleByHeader("Status").SetBoldWhenTextEquals("ready")
            .SetBoldByTextSet(new HashSet<string> { "READY" }, bold: false)
            .SetFillColorWhenTextEquals("PENDING", OfficeColor.Accent3)
            .SetColorByTextMap(new Dictionary<string, OfficeColor> { ["READY"] = OfficeColor.Accent4 });
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        WorksheetPart worksheet = package.WorkbookPart!.WorksheetParts.Single();
        Stylesheet styles = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
        Assert.False(Font(styles, Format(worksheet, styles, "A1")).Bold?.Val?.Value ?? false);
        Assert.Equal(0U, Format(worksheet, styles, "A1").NumberFormatId!.Value);
        foreach (string reference in new[] { "A2", "A3" }) {
            CellFormat format = Format(worksheet, styles, reference);
            Font font = Font(styles, format);
            Assert.False(font.Bold?.Val?.Value ?? font.Bold != null);
            Assert.False(font.Italic?.Val?.Value ?? font.Italic != null);
            Assert.Equal(UnderlineValues.None, font.Underline?.Val?.Value ?? UnderlineValues.None);
            Assert.False(font.Strike?.Val?.Value ?? font.Strike != null);
            Assert.Equal("Aptos", font.FontName!.Val!.Value);
            Assert.Equal(14D, font.FontSize!.Val!.Value);
            Assert.Equal(HorizontalAlignmentValues.Right, format.Alignment!.Horizontal!.Value);
            Assert.Equal(VerticalAlignmentValues.Center, format.Alignment.Vertical!.Value);
            Assert.False(format.Alignment.WrapText?.Value ?? false);
            Assert.Equal(VerticalAlignmentRunValues.Superscript, font.VerticalTextAlignment!.Val!.Value);
        }
        AssertTheme(Font(styles, Format(worksheet, styles, "A2")).Color!, 7U);
        AssertTheme(Fill(styles, Format(worksheet, styles, "A3")).PatternFill!.ForegroundColor!, 6U);
    }

    [Fact]
    public void TypedColorsPersistNativeSlotsAndArgbWithoutLosingFontClassification() {
        OfficeColor translucent = OfficeColor.FromRgba(0x12, 0x34, 0x56, 0x80);
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(success: translucent, accent4: OfficeColor.Parse("#ABCDEF")));
        using var document = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme });
        ExcelSheet sheet = document.AddSheet("Colors");
        sheet.Cell(1, 1).SetValue("Text").SetColor(OfficeColor.Accent1).SetFillColor(OfficeColor.Accent2)
            .SetBorder(ExcelBorderStyle.Thin, OfficeColor.Accent3).SetDiagonalBorder(ExcelBorderStyle.Dotted, OfficeColor.Accent4);
        sheet.Range("B1:B2").SetColor(translucent).SetFillColor(OfficeColor.Success)
            .SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Double, OfficeColor.Accent5)
            .SetGradientFill(OfficeColor.Accent6, translucent);
        sheet.SetRowLayout(3, new ExcelRowLayoutOptions { FirstColumn = 1, LastColumn = 1, FontFamily = "Courier New", BackgroundColor = OfficeColor.Accent5 });
        document.DefineNamedStyle("Themed", new ExcelStyleDefinition { FontFamily = "Arial", Color = OfficeColor.Accent6, BackgroundColor = OfficeColor.Success });
        sheet.Cell(4, 1).SetValue("Named");
        sheet.ApplyNamedStyle("Themed", "A4:A4");
        var rich = new ExcelRichTextRun("Rich") { FontFamily = "Aptos", FontFamilyClassification = 2, Color = OfficeColor.Accent4, Bold = false };
        sheet.Cell(5, 1).SetRichText(rich, new ExcelRichTextRun(" Literal") { Color = translucent });
        sheet.SetCommentRichText(5, 1, new[] { rich });
        sheet.Range("D1:D2").ConditionalFormat().CellIs(ExcelConditionalFormattingOperator.GreaterThan, "0", fillColor: OfficeColor.Accent2);
        sheet.Range("E1:E2").ConditionalFormat().Formula("E1>0", fillColor: OfficeColor.Accent3);
        sheet.Range("F1:F2").ConditionalFormat().ColorScale(OfficeColor.Accent1, OfficeColor.Accent3, OfficeColor.Accent6);
        sheet.Range("G1:G2").ConditionalFormat().DataBar(OfficeColor.Success);
        byte[] bytes = document.ToBytes();
        using (var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            WorksheetPart worksheet = package.WorkbookPart!.WorksheetParts.Single();
            Stylesheet styles = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            CellFormat a1 = Format(worksheet, styles, "A1");
            AssertTheme(Font(styles, a1).Color!, 4U);
            AssertTheme(Fill(styles, a1).PatternFill!.ForegroundColor!, 5U);
            Border border = styles.Borders!.Elements<Border>().ElementAt((int)a1.BorderId!.Value);
            AssertTheme(border.LeftBorder!.Color!, 6U);
            AssertTheme(border.DiagonalBorder!.Color!, 7U);
            foreach (string reference in new[] { "B1", "B2" }) {
                CellFormat format = Format(worksheet, styles, reference);
                Assert.Equal("80123456", Font(styles, format).Color!.Rgb!.Value);
                GradientStop[] stops = Fill(styles, format).GradientFill!.Elements<GradientStop>().ToArray();
                AssertTheme(stops[0].Color!, 9U);
                Assert.Equal("80123456", stops[1].Color!.Rgb!.Value);
                Border ranged = styles.Borders.Elements<Border>().ElementAt((int)format.BorderId!.Value);
                AssertTheme(ranged.BottomBorder!.Color!, 8U);
            }
            Assert.Equal("Courier New", Font(styles, Format(worksheet, styles, "A3")).FontName!.Val!.Value);
            AssertTheme(Fill(styles, Format(worksheet, styles, "A3")).PatternFill!.ForegroundColor!, 8U);
            AssertTheme(Font(styles, Format(worksheet, styles, "A4")).Color!, 9U);
            Assert.Equal("80123456", Fill(styles, Format(worksheet, styles, "A4")).PatternFill!.ForegroundColor!.Rgb!.Value);
            Run[] runs = FindCell(worksheet, "A5").InlineString!.Elements<Run>().ToArray();
            Assert.Equal("Aptos", runs[0].RunProperties!.GetFirstChild<RunFont>()!.Val!.Value);
            Assert.Equal(2, runs[0].RunProperties!.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.FontFamily>()!.Val!.Value);
            Assert.False(runs[0].RunProperties!.GetFirstChild<Bold>()!.Val!.Value);
            AssertTheme(runs[0].RunProperties!.GetFirstChild<Color>()!, 7U);
            Assert.Equal("80123456", runs[1].RunProperties!.GetFirstChild<Color>()!.Rgb!.Value);
            AssertTheme(worksheet.WorksheetCommentsPart!.Comments.Descendants<RunProperties>().Single().GetFirstChild<Color>()!, 7U);
            Color[] fills = styles.DifferentialFormats!.Descendants<ForegroundColor>().Select(c => new Color { Theme = c.Theme, Rgb = c.Rgb }).ToArray();
            Assert.Equal(new uint?[] { 5U, 6U }, fills.Select(c => c.Theme?.Value));
            Assert.Equal("80123456", worksheet.Worksheet.Descendants<DataBar>().Single().Elements<Color>().Single().Rgb!.Value);
            ColorScale scale = worksheet.Worksheet.Descendants<ColorScale>().Single();
            Assert.Equal(new uint?[] { 4U, 6U, 9U }, scale.Elements<Color>().Select(c => c.Theme?.Value));
            ConditionalFormatValueObject[] thresholds = scale.Elements<ConditionalFormatValueObject>().ToArray();
            Assert.Equal(ConditionalFormatValueObjectValues.Percentile, thresholds[1].Type!.Value);
            Assert.Equal("50", thresholds[1].Val!.Value);
        }
        using var reopened = ExcelDocument.Load(new MemoryStream(bytes));
        ExcelRichTextRun[] loaded = reopened.Sheets[0].Cell(5, 1).GetRichText().ToArray();
        Assert.Equal("Aptos", loaded[0].FontFamily);
        Assert.Equal((byte)2, loaded[0].FontFamilyClassification);
        Assert.Equal(OfficeColor.Accent4, loaded[0].Color);
        Assert.Equal(translucent, loaded[1].Color);
        ExcelVisualTextRun[] visual = reopened.Sheets[0].Range("A5:A5").CreateVisualSnapshot().Cells.Single().RichTextRuns.ToArray();
        Assert.Equal("FFABCDEF", visual[0].FontColorArgb);
        Assert.Equal("80123456", visual[1].FontColorArgb);
        Assert.Equal("Aptos", visual[0].FontName);
        using var xmlCopy = ExcelDocument.ImportSpreadsheetXml(reopened.ExportSpreadsheetXml().Value).Value;
        ExcelRichTextRun[] xmlRuns = xmlCopy.Sheets[0].Cell(5, 1).GetRichText().ToArray();
        Assert.Equal(OfficeColor.Parse("#ABCDEF"), xmlRuns[0].Color);
        Assert.Equal("Aptos", xmlRuns[0].FontFamily);
        Assert.Equal(OfficeColor.Parse("#123456"), xmlRuns[1].Color);
        using var legacySource = ExcelDocument.Create();
        legacySource.AddSheet("Legacy").Cell(1, 1).SetRichText(new ExcelRichTextRun("Native") {
            FontFamily = "Aptos", FontFamilyClassification = 2, Color = OfficeColor.Parse("#ABCDEF")
        });
        using var xls = new MemoryStream();
        legacySource.Save(xls, ExcelFileFormat.Xls);
        xls.Position = 0;
        using var legacyCopy = ExcelDocument.Load(xls);
        ExcelRichTextRun[] legacyRuns = legacyCopy.Sheets[0].Cell(1, 1).GetRichText().ToArray();
        Assert.Equal("Aptos", legacyRuns[0].FontFamily);
        Assert.Equal((byte)2, legacyRuns[0].FontFamilyClassification);
        Assert.Equal(OfficeColor.Parse("#ABCDEF"), legacyRuns[0].Color);
    }

    [Fact]
    public void PartialChartUpdatesPreserveOmittedGeometryFlagsAndSymbols() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Chart");
        var chart = sheet.AddChart(OfficeChartType.Line, new ExcelChartData(new[] { "A", "B" },
            new[] { new ExcelChartSeries("Revenue", new[] { 1D, 2D }) }), new ExcelChartOptions { Row = 1, Column = 3 });
        chart.SetLegend(OfficeChartLegendPosition.Left, overlay: true)
            .SetDataLabels(true, true, true, true, true, OfficeChartDataLabelPosition.Right, "0.00", true)
            .SetCategoryAxisGridlines(true, true).SetValueAxisGridlines(true, true)
            .SetSeriesMarker(0, OfficeChartMarkerShape.Diamond, OfficeLength.Points(9));
        chart.SetLegendTextStyle(color: OfficeColor.Accent2).SetDataLabelTextStyle(color: OfficeColor.Accent3)
            .SetDataLabels(showValue: false).SetDataLabels(numberFormat: "0.000")
            .SetCategoryAxisGridlines(lineColor: OfficeColor.Red).SetValueAxisGridlines(lineColor: OfficeColor.Red)
            .SetSeriesMarker(0, fillColor: OfficeColor.Green);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false)) {
            C.Chart native = package.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single().ChartSpace.GetFirstChild<C.Chart>()!;
            C.Legend legend = native.GetFirstChild<C.Legend>()!;
            Assert.Equal(C.LegendPositionValues.Left, legend.GetFirstChild<C.LegendPosition>()!.Val!.Value);
            Assert.True(legend.GetFirstChild<C.Overlay>()!.Val!.Value);
            C.DataLabels labels = native.Descendants<C.LineChart>().Single().GetFirstChild<C.DataLabels>()!;
            Assert.False(labels.GetFirstChild<C.ShowValue>()!.Val!.Value);
            Assert.True(labels.GetFirstChild<C.ShowCategoryName>()!.Val!.Value);
            Assert.True(labels.GetFirstChild<C.ShowSeriesName>()!.Val!.Value);
            Assert.True(labels.GetFirstChild<C.ShowLegendKey>()!.Val!.Value);
            Assert.True(labels.GetFirstChild<C.ShowPercent>()!.Val!.Value);
            Assert.Equal(C.DataLabelPositionValues.Right, labels.GetFirstChild<C.DataLabelPosition>()!.Val!.Value);
            Assert.Equal("0.000", labels.GetFirstChild<C.NumberingFormat>()!.FormatCode!.Value);
            Assert.True(labels.GetFirstChild<C.NumberingFormat>()!.SourceLinked!.Value);
            Assert.Equal(2, native.Descendants<C.MajorGridlines>().Count());
            Assert.Equal(2, native.Descendants<C.MinorGridlines>().Count());
            C.Marker marker = native.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Marker>()!;
            Assert.Equal(C.MarkerStyleValues.Diamond, marker.Symbol!.Val!.Value);
            Assert.Equal((byte)9, marker.Size!.Val!.Value);
        }
        chart.SetLegend(overlay: false).SetDataLabels(sourceLinked: false)
            .SetCategoryAxisGridlines(showMajor: false).SetValueAxisGridlines(showMinor: false)
            .SetSeriesMarkerByName("Revenue", size: OfficeLength.Points(11));
        using (var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false)) {
            C.Chart native = package.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single().ChartSpace.GetFirstChild<C.Chart>()!;
            Assert.Equal(C.LegendPositionValues.Left, native.GetFirstChild<C.Legend>()!.GetFirstChild<C.LegendPosition>()!.Val!.Value);
            Assert.False(native.GetFirstChild<C.Legend>()!.GetFirstChild<C.Overlay>()!.Val!.Value);
            C.NumberingFormat format = native.Descendants<C.LineChart>().Single().GetFirstChild<C.DataLabels>()!.GetFirstChild<C.NumberingFormat>()!;
            Assert.Equal("0.000", format.FormatCode!.Value);
            Assert.False(format.SourceLinked!.Value);
            Assert.Null(native.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.MajorGridlines>());
            Assert.NotNull(native.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.MinorGridlines>());
            Assert.NotNull(native.Descendants<C.ValueAxis>().Single().GetFirstChild<C.MajorGridlines>());
            Assert.Null(native.Descendants<C.ValueAxis>().Single().GetFirstChild<C.MinorGridlines>());
            C.Marker marker = native.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Marker>()!;
            Assert.Equal(C.MarkerStyleValues.Diamond, marker.Symbol!.Val!.Value);
            Assert.Equal((byte)11, marker.Size!.Val!.Value);
        }
        chart.HideLegend().SetLegend();
        using var restored = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        Assert.Single(restored.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single().ChartSpace.Descendants<C.Legend>());
    }

    [Theory]
    [InlineData(ExcelConditionalFormattingSource.Standard)]
    [InlineData(ExcelConditionalFormattingSource.Office2010Extension)]
    public void ConditionalColorUpdatesRetainOtherStyleAndKeepThemeLinks(ExcelConditionalFormattingSource source) {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(accent4: OfficeColor.Parse("#A1B2C3"), danger: OfficeColor.Parse("#1122AA")))
            .WithSurfaces(OfficeTheme.Modern.Surfaces.With(dangerFill: OfficeColor.Parse("#FFEEDD")));
        using var document = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme });
        var sheet = document.AddSheet("Rules");
        var rule = sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
            Source = source, Range = "A1:A3", Type = "Expression", Formulas = new[] { "A1>0" },
            DifferentialFontBold = true, DifferentialFontColorArgb = "FF112233", DifferentialFillColorArgb = "FF445566"
        });
        var sibling = sheet.CloneConditionalFormattingRule(rule, "B1:B3");
        rule.StopIfTrue = true;
        sheet.UpdateConditionalFormattingRule(rule);
        var updated = sheet.SetConditionalFormattingColors(rule, OfficeColor.Accent4, OfficeColor.FromRgba(0x12, 0x34, 0x56, 0x80));
        Assert.True(updated.StopIfTrue);
        Assert.True(updated.DifferentialFontBold);
        Assert.Equal("FFA1B2C3", updated.DifferentialFontColorArgb);
        Assert.Equal("80123456", updated.DifferentialFillColorArgb);
        Assert.Equal("FF112233", sheet.GetConditionalFormattingRules("B1:B3").Single().DifferentialFontColorArgb);
        var defaults = sheet.SetConditionalFormattingColors(sibling);
        Assert.Equal("FF1122AA", defaults.DifferentialFontColorArgb);
        Assert.Equal("FFFFEEDD", defaults.DifferentialFillColorArgb);
        Assert.True(defaults.DifferentialFontBold);
        Assert.Throws<InvalidOperationException>(() => document.AddSheet("Other").SetConditionalFormattingColors(rule, OfficeColor.Red));
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        WorkbookPart workbook = package.WorkbookPart!;
        OpenXmlCompositeElement differential = source == ExcelConditionalFormattingSource.Standard
            ? workbook.WorkbookStylesPart!.Stylesheet.DifferentialFormats!.Elements<DifferentialFormat>().ElementAt((int)updated.DifferentialFormatId!.Value)
            : workbook.WorksheetParts.SelectMany(part => part.Worksheet.Descendants<X14.DifferentialType>()).First();
        Assert.True(differential.GetFirstChild<Font>()!.Bold!.Val!.Value);
        AssertTheme(differential.GetFirstChild<Font>()!.Color!, 7U);
        Assert.Equal("80123456", differential.GetFirstChild<Fill>()!.PatternFill!.ForegroundColor!.Rgb!.Value);
    }

    [Fact]
    public void SharedChartColorsKeepRgbAlphaAndNativeThemeRoles() {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(danger: OfficeColor.Parse("#987654")));
        using var document = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme });
        var sheet = document.AddSheet("Chart colors");
        var translucent = OfficeColor.FromRgba(0x12, 0x34, 0x56, 0x80);
        var series = new OfficeSeries("Color", new[] { 1D, 2D, 3D }, null, translucent,
            new OfficeColor?[] { OfficeColor.Accent4, translucent, OfficeColor.Danger }, showMarkers: true,
            markerOutlineColor: OfficeColor.Accent2);
        sheet.AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] { series }));
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        C.LineChartSeries native = package.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace.Descendants<C.LineChartSeries>().Single();
        A.RgbColorModelHex fill = native.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!;
        Assert.Equal("123456", fill.Val!.Value);
        Assert.Equal(50196, fill.GetFirstChild<A.Alpha>()!.Val!.Value);
        C.ChartShapeProperties marker = native.GetFirstChild<C.Marker>()!.GetFirstChild<C.ChartShapeProperties>()!;
        Assert.Equal("123456", marker.GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        Assert.Equal(A.SchemeColorValues.Accent2, marker.GetFirstChild<A.Outline>()!.Descendants<A.SchemeColor>().Single().Val!.Value);
        C.DataPoint[] points = native.Elements<C.DataPoint>().ToArray();
        Assert.Equal(A.SchemeColorValues.Accent4, points[0].Descendants<A.SchemeColor>().Single().Val!.Value);
        Assert.Equal("123456", points[1].Descendants<A.RgbColorModelHex>().Single().Val!.Value);
        Assert.Equal(50196, points[1].Descendants<A.Alpha>().Single().Val!.Value);
        Assert.Equal("987654", points[2].Descendants<A.RgbColorModelHex>().Single().Val!.Value);
    }

    private static Cell FindCell(WorksheetPart worksheet, string reference) =>
        worksheet.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == reference);

    private static CellFormat Format(WorksheetPart worksheet, Stylesheet styles, string reference) =>
        styles.CellFormats!.Elements<CellFormat>().ElementAt((int)(FindCell(worksheet, reference).StyleIndex?.Value ?? 0U));

    private static Font Font(Stylesheet styles, CellFormat format) => styles.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
    private static Fill Fill(Stylesheet styles, CellFormat format) => styles.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value);

    private static void AssertTheme(ColorType color, uint index) {
        Assert.Equal(index, color.Theme!.Value);
        Assert.Null(color.Rgb);
    }
}
