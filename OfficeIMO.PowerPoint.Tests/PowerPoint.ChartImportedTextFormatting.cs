using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class PowerPointChartImportedTextFormattingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ImportedChartStylePreservesScriptFontsAndSchemaOrder(bool legend, bool replaceFont) {
        using var source = PowerPointPresentation.Create();
        PowerPointChart chart = source.AddSlide().AddChart().SetTitle("Imported title")
            .SetLegend(OfficeChartLegendPosition.Right)
            .SetTitleTextStyle(fontName: "Aptos")
            .SetLegendTextStyle(fontName: "Aptos");
        C.Chart chartElement = source.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .ChartParts.Single().ChartSpace.GetFirstChild<C.Chart>()!;
        A.TextCharacterPropertiesType properties = legend
            ? chartElement.GetFirstChild<C.Legend>()!.Descendants<A.DefaultRunProperties>().Single()
            : chartElement.GetFirstChild<C.Title>()!.Descendants<A.RunProperties>().Single();
        properties.AddChild(new A.EastAsianFont { Typeface = "Noto Sans CJK JP" }, true);
        properties.AddChild(new A.ComplexScriptFont { Typeface = "Noto Sans Arabic" }, true);
        Assert.Empty(new OpenXmlValidator().Validate(source.OpenXmlDocument));
        using var input = new MemoryStream();
        source.Save(input);
        input.Position = 0;
        using var imported = PowerPointPresentation.Load(input);
        PowerPointChart actualChart = imported.Slides.Single().Charts.Single();
        if (legend) actualChart.SetLegendTextStyle(color: "123456", fontName: replaceFont ? "Arial" : null);
        else actualChart.SetTitleTextStyle(color: "123456", fontName: replaceFont ? "Arial" : null);
        using var output = new MemoryStream();
        imported.Save(output);
        output.Position = 0;
        using var reopened = PowerPointPresentation.Load(output);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        C.Chart finalChart = reopened.OpenXmlDocument.PresentationPart!.SlideParts.Single()
            .ChartParts.Single().ChartSpace.GetFirstChild<C.Chart>()!;
        A.TextCharacterPropertiesType actual = legend
            ? finalChart.GetFirstChild<C.Legend>()!.Descendants<A.DefaultRunProperties>().Single()
            : finalChart.GetFirstChild<C.Title>()!.Descendants<A.RunProperties>().Single();
        Assert.Equal(replaceFont ? "Arial" : "Aptos", actual.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
        Assert.Equal("Noto Sans CJK JP", actual.GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
        Assert.Equal("Noto Sans Arabic", actual.GetFirstChild<A.ComplexScriptFont>()!.Typeface!.Value);
        Assert.Equal("123456", actual.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Equal("Imported title", finalChart.GetFirstChild<C.Title>()!.Descendants<A.Text>().Single().Text);
    }
}
