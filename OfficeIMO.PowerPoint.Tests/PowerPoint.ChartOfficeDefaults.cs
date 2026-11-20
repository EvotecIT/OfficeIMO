using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

/// <summary>Charts written with the formatting Excel, Word and PowerPoint apply by default.</summary>
public sealed class PowerPointChartOfficeDefaultsTests {
    private const string Accent1Line = "<c:spPr><a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>";
    private const string ExcelChartArea = "<c:spPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>";
    private const string SeriesLabels = "<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>" +
        "<c:txPr><a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" lIns=\"38100\" tIns=\"19050\" rIns=\"38100\" bIns=\"19050\" anchor=\"ctr\" anchorCtr=\"1\"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/></a:schemeClr></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>" +
        "<c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/>" +
        "<c:extLst><c:ext uri=\"{CE6537A1-D6FC-4f65-9D91-7224C49458BB}\" xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\"><c15:showLeaderLines val=\"1\"/><c15:leaderLines><c:spPr><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:ln></c:spPr></c15:leaderLines></c:ext></c:extLst></c:dLbls>";

    [Fact]
    public void LineSeriesWithRoundCapsAreDrawnWithRoundCaps() {
        string svg = RenderSvg(CategoryChart("lineChart", "<c:grouping val=\"standard\"/>", Accent1Line + "<c:marker><c:symbol val=\"none\"/></c:marker>"),
            OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains("stroke-linecap=\"round\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void PieWithOfficeChartAreaOutlineIsDrawn() {
        string plot = "<c:pieChart><c:varyColors val=\"1\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" + SeriesName() + Categories() + Values() +
            "</c:ser><c:firstSliceAng val=\"0\"/></c:pieChart>";

        RenderSvg(ChartSpace(plot, ExcelChartArea), OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
    }

    [Theory]
    [InlineData("#,##0;[Red]#,##0", "4")]
    [InlineData("#,##0;-#,##0;\\-", ">-<")]
    [InlineData("\\$#,##0;[Red]\\(\\$#,##0\\);\\-", "$4")]
    [InlineData("[$€-413] #,##0", "€ 4")]
    [InlineData("0\";kg\";(0);\"-\"", ">2;kg<")]
    [InlineData("[>=45000]yyyy-mm-dd;\"early\"", ">early<")]
    public void OfficeAxisNumberFormatsAreDrawn(string format, string expectedText) {
        string svg = RenderSvg(CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", string.Empty, format),
            OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation);
        Assert.Contains(expectedText, svg, StringComparison.Ordinal);
    }

    [Fact]
    public void InvalidAxisNumberFormatIsDrawnAsGeneralWithADiagnostic() {
        string svg = RenderSvg(CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", string.Empty, "#,##0\"x"),
            OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.Message.IndexOf("drawn as General", StringComparison.Ordinal) >= 0);
        Assert.Contains(">4<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void SeriesDataLabelsAreDrawnAndDeletedPointLabelsStayHidden() {
        string labels = SeriesLabels.Replace("<c:dLbls>", "<c:dLbls><c:dLbl><c:idx val=\"2\"/><c:delete val=\"1\"/></c:dLbl>");
        string svg = RenderSvg(CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>",
                "<c:spPr><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></c:spPr><c:invertIfNegative val=\"0\"/>" + labels),
            OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(">4.3<", svg, StringComparison.Ordinal);
        Assert.Contains(">4.5<", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(">3.5<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void BestEffortDrawsAChartThatStrictSkipsAndReportsTheApproximation() {
        string errorBars = "<c:spPr><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></c:spPr><c:invertIfNegative val=\"0\"/>" +
            "<c:errBars><c:errBarType val=\"both\"/><c:errValType val=\"stdErr\"/><c:noEndCap val=\"0\"/></c:errBars>";
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", errorBars);

        RenderSvg(chart, OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> strict);
        string svg = RenderSvg(chart, OfficeChartFidelity.BestEffort, out IReadOnlyList<OfficeImageExportDiagnostic> bestEffort);

        Assert.Contains(strict, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.UnsupportedChart &&
            diagnostic.Message.IndexOf("Error bars", StringComparison.Ordinal) >= 0);
        AssertDrawn(bestEffort);
        OfficeImageExportDiagnostic approximation = Assert.Single(bestEffort, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation);
        Assert.Equal("Error bars are left out.", approximation.Message);
        Assert.Equal(OfficeConversionLossKind.Approximation, approximation.LossKind);
        Assert.Contains(">North<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfExportHonoursBestEffortChartFidelity() {
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>",
            "<c:spPr><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:ln cmpd=\"dbl\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:ln></c:spPr>");
        using MemoryStream stream = CreateDeck(chart);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);

        PdfDocumentConversionResult strict = presentation.ConvertToPdfWithEvidence();
        PdfDocumentConversionResult bestEffort = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions {
            ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort }
        });

        Assert.Contains(strict.Warnings, warning => warning.Code == "unsupported-chart");
        Assert.DoesNotContain(bestEffort.Warnings, warning => warning.Code == "unsupported-chart");
        Assert.Contains(bestEffort.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    [Fact]
    public void StrictSkipsAChartWhoseSurfacesCannotBeDrawnAndBestEffortReportsThem() {
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", string.Empty)
            .Replace("<c:roundedCorners val=\"0\"/>", "<c:roundedCorners val=\"1\"/>");

        RenderSvg(chart, OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> strict);
        RenderSvg(chart, OfficeChartFidelity.BestEffort, out IReadOnlyList<OfficeImageExportDiagnostic> bestEffort);

        Assert.Contains(strict, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.UnsupportedChart);
        AssertDrawn(bestEffort);
        Assert.Contains(bestEffort, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.Message == "The rounded chart frame is drawn with square corners.");
    }

    [Fact]
    public void SingleChartImageHonoursTheChartFidelity() {
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", string.Empty)
            .Replace("<c:roundedCorners val=\"0\"/>", "<c:roundedCorners val=\"1\"/>");
        using MemoryStream stream = CreateDeck(chart);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        PowerPointChart native = presentation.Slides[0].Charts.Single();

        Assert.Throws<NotSupportedException>(() => native.ToImage(new OfficeChartImageExportOptions { Format = OfficeImageExportFormat.Svg }));
        OfficeImageExportResult image = native.ToImage(new OfficeChartImageExportOptions {
            Format = OfficeImageExportFormat.Svg,
            ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort }
        });

        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.Message == "The rounded chart frame is drawn with square corners.");
    }

    [Fact]
    public void UnresolvedWorkbookLinkedAxisFormatIsReported() {
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", string.Empty)
            .Replace("sourceLinked=\"0\"", "sourceLinked=\"1\"").Replace("<c:formatCode>General</c:formatCode>", "<c:formatCode>0.0</c:formatCode>")
            .Replace("</c:chartSpace>", "<c:externalData r:id=\"rIdMissing\"><c:autoUpdate val=\"0\"/></c:externalData></c:chartSpace>");

        RenderSvg(chart, OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.Message.IndexOf("could not be resolved", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void SeriesLabelPositionOverridesTheGroupPositionBeforeItIsChecked() {
        string seriesLabels = "<c:dLbls><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbls>";
        string plot = "<c:pieChart><c:varyColors val=\"1\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" + SeriesName() + seriesLabels +
            Categories() + Values() + "</c:ser><c:dLbls><c:dLblPos val=\"inEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbls>" +
            "<c:firstSliceAng val=\"0\"/></c:pieChart>";

        string svg = RenderSvg(ChartSpace(plot, ExcelChartArea), OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(">4.3<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void PointLabelDeletedOnTheChartGroupStaysHidden() {
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", string.Empty)
            .Replace("<c:gapWidth", "<c:dLbls><c:dLbl><c:idx val=\"0\"/><c:delete val=\"1\"/></c:dLbl><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbls><c:gapWidth");

        string svg = RenderSvg(chart, OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(">2.5<", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(">4.3<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void BestEffortDrawsAPointWithMixedCustomTextAsAStandardLabel() {
        string labels = SeriesLabels.Replace("<c:dLbls>", "<c:dLbls><c:dLbl><c:idx val=\"1\"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" +
            "<a:r><a:rPr lang=\"en-US\" b=\"1\"/><a:t>Bold</a:t></a:r><a:r><a:rPr lang=\"en-US\" b=\"0\"/><a:t> plain</a:t></a:r></a:p></c:rich></c:tx>" +
            "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbl>");
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", labels);

        string svg = RenderSvg(chart, OfficeChartFidelity.BestEffort, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.Message == "A data label with its own text or position is drawn as a standard label.");
        Assert.Contains(">2.5<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void PointLabelKeepsTheFormatOfAHiddenGroup() {
        string seriesLabels = "<c:dLbls><c:dLbl><c:idx val=\"0\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbl>" +
            "<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/></c:dLbls>";
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>", seriesLabels)
            .Replace("<c:gapWidth", "<c:dLbls><c:numFmt formatCode=\"0.00\" sourceLinked=\"0\"/><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/></c:dLbls><c:gapWidth");

        string svg = RenderSvg(chart, OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains(">4.30<", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(">2.50<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void SeriesLabelTextKeepsTheGroupColourWhenItOnlySetsTheSize() {
        string seriesLabels = "<c:dLbls>" + TextProperties("<a:defRPr sz=\"1200\"/>") + "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbls>";
        string chart = CategoryChart("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>",
                "<c:spPr><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></c:spPr><c:invertIfNegative val=\"0\"/>" + seriesLabels)
            .Replace("<c:gapWidth", "<c:dLbls>" + TextProperties("<a:defRPr><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill></a:defRPr>") +
                "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbls><c:gapWidth")
            .Replace("</c:chart>" + ExcelChartArea, "</c:chart>" + ExcelChartArea +
                TextProperties("<a:defRPr><a:solidFill><a:srgbClr val=\"0070C0\"/></a:solidFill></a:defRPr>"));

        string svg = RenderSvg(chart, OfficeChartFidelity.BestEffort, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.Contains("c00000", svg.ToLowerInvariant(), StringComparison.Ordinal);
    }

    [Fact]
    public void MovingAverageLongerThanDrawingMlAllowsIsNotDrawn() {
        string chart = CategoryChart("lineChart", "<c:grouping val=\"standard\"/>", Accent1Line + "<c:marker><c:symbol val=\"none\"/></c:marker>" +
            "<c:trendline><c:trendlineType val=\"movingAvg\"/><c:period val=\"300\"/><c:dispRSqr val=\"0\"/><c:dispEq val=\"0\"/></c:trendline>");

        RenderSvg(chart, OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> strict);
        RenderSvg(chart, OfficeChartFidelity.BestEffort, out IReadOnlyList<OfficeImageExportDiagnostic> bestEffort);

        Assert.Contains(strict, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.UnsupportedChart &&
            diagnostic.Message.IndexOf("trendline type", StringComparison.Ordinal) >= 0);
        AssertDrawn(bestEffort);
        Assert.Contains(bestEffort, diagnostic => diagnostic.Message == "A trendline the renderer cannot fit is left out.");
    }

    [Fact]
    public void TrendlineWithoutALineOrLegendEntryDrawsOnlyItsEquation() {
        string series = Accent1Line + "<c:marker><c:symbol val=\"none\"/></c:marker>";
        string trendline = "<c:trendline><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:trendlineType val=\"linear\"/><c:dispRSqr val=\"0\"/><c:dispEq val=\"1\"/></c:trendline>";
        string Chart(string formatting) => CategoryChart("lineChart", "<c:grouping val=\"standard\"/>", formatting)
            .Replace("<c:legendPos val=\"b\"/>", "<c:legendPos val=\"b\"/><c:legendEntry><c:idx val=\"1\"/><c:delete val=\"1\"/></c:legendEntry>");

        string without = RenderSvg(Chart(series), OfficeChartFidelity.Strict, out _);
        string svg = RenderSvg(Chart(series + trendline), OfficeChartFidelity.Strict, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics);

        AssertDrawn(diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation);
        Assert.Contains("y = ", svg, StringComparison.Ordinal);
        Assert.Equal(CountPaths(without), CountPaths(svg));
    }

    [Fact]
    public void BestEffortKeepsBubblesWhoseErrorBarsCannotBeDrawn() {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddChart(OfficeChartType.Bubble, new OfficeChartData(new[] { "A", "B" }, new[] {
            OfficeSeries.CreateBubble("Values", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 6d })
        }));
        C.BubbleChartSeries series = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.BubbleChartSeries>().Single();
        series.InsertBefore(new C.ErrorBars(new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y },
            new C.ErrorBarType { Val = C.ErrorBarValues.Both }, new C.ErrorBarValueType { Val = C.ErrorValues.StandardError },
            new C.NoEndCap { Val = false }), series.GetFirstChild<C.XValues>());

        IReadOnlyList<OfficeImageExportDiagnostic> Render(OfficeChartFidelity fidelity) => slide.ToImage(new PowerPointImageExportOptions {
            Format = OfficeImageExportFormat.Svg,
            ChartRendering = { Fidelity = fidelity }
        }).Diagnostics;

        Assert.Contains(Render(OfficeChartFidelity.Strict), diagnostic => diagnostic.Code == OfficeDiagnosticCodes.UnsupportedChart);
        IReadOnlyList<OfficeImageExportDiagnostic> bestEffort = Render(OfficeChartFidelity.BestEffort);
        AssertDrawn(bestEffort);
        Assert.Contains(bestEffort, diagnostic => diagnostic.Message == "Error bars are left out.");
    }

    private static int CountPaths(string svg) => System.Text.RegularExpressions.Regex.Matches(svg, "<path").Count;

    private static string TextProperties(string runProperties) =>
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>" + runProperties + "</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>";

    private static void AssertDrawn(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) =>
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.UnsupportedChart);

    private static string RenderSvg(string chartXml, OfficeChartFidelity fidelity, out IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
        using MemoryStream stream = CreateDeck(chartXml);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        OfficeImageExportResult image = presentation.Slides[0].ToImage(new PowerPointImageExportOptions {
            Format = OfficeImageExportFormat.Svg,
            ChartRendering = { Fidelity = fidelity }
        });
        diagnostics = image.Diagnostics;
        return Encoding.UTF8.GetString(image.Bytes);
    }

    private static MemoryStream CreateDeck(string chartXml) {
        var stream = new MemoryStream();
        using (PowerPointDocument deck = PowerPointDocument.Create(stream, new PowerPointCreateOptions())) {
            PowerPointSlide slide = deck.AddSlide(PowerPointSlideLayoutType.Blank, 0);
            slide.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "a" }, new[] { new OfficeSeries("Value", new[] { 1d }) }));
            deck.Save();
        }
        stream.Position = 0;
        using (PresentationDocument package = PresentationDocument.Open(stream, true)) {
            SlidePart slidePart = package.PresentationPart!.SlideParts.First();
            P.GraphicFrame frame = slidePart.Slide!.Descendants<P.GraphicFrame>().First();
            frame.Transform = new P.Transform(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 5486400, Cy = 3200400 });
            using (Stream part = slidePart.ChartParts.First().GetStream(FileMode.Create))
            using (var writer = new StreamWriter(part, new UTF8Encoding(false))) writer.Write(chartXml);
            slidePart.Slide.Save();
        }
        stream.Position = 0;
        return stream;
    }

    private static string CategoryChart(string group, string groupHeader, string seriesFormatting, string valueFormat = "General") =>
        ChartSpace("<c:" + group + ">" + groupHeader + "<c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" + SeriesName() +
            seriesFormatting + Categories() + Values() + (group == "lineChart" ? "<c:smooth val=\"0\"/>" : string.Empty) + "</c:ser>" +
            (group == "lineChart" ? "<c:marker val=\"1\"/>" : "<c:gapWidth val=\"219\"/>") +
            "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:" + group + ">" +
            "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/></c:catAx>" +
            "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:numFmt formatCode=\"" +
            System.Security.SecurityElement.Escape(valueFormat) + "\" sourceLinked=\"0\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>",
            ExcelChartArea);

    private static string ChartSpace(string plotArea, string chartArea) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
        "<c:roundedCorners val=\"0\"/><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>" + plotArea + "</c:plotArea>" +
        "<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"0\"/></c:legend><c:plotVisOnly val=\"1\"/></c:chart>" + chartArea + "</c:chartSpace>";

    private static string SeriesName() =>
        "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>";

    private static string Categories() =>
        "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f><c:strCache><c:ptCount val=\"4\"/><c:pt idx=\"0\"><c:v>North</c:v></c:pt><c:pt idx=\"1\"><c:v>South</c:v></c:pt><c:pt idx=\"2\"><c:v>East</c:v></c:pt><c:pt idx=\"3\"><c:v>West</c:v></c:pt></c:strCache></c:strRef></c:cat>";

    private static string Values() =>
        "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"4\"/><c:pt idx=\"0\"><c:v>4.3</c:v></c:pt><c:pt idx=\"1\"><c:v>2.5</c:v></c:pt><c:pt idx=\"2\"><c:v>3.5</c:v></c:pt><c:pt idx=\"3\"><c:v>4.5</c:v></c:pt></c:numCache></c:numRef></c:val>";
}
