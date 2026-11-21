using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

// .NET Framework's general number format keeps 15 significant digits; these parts carry the shortest round-trip
// text that .NET 8 and later write, on every framework.
public sealed class WordNumberTextTests {
    [Fact]
    public void ChartCachesShapeStylesAndCustomPropertiesWriteTheSameNumberTextOnEveryFramework() {
        using var document = WordDocument.Create();
        var chart = document.AddChart("Thirds");
        chart.AddCategories(new[] { "A" });
        chart.AddBar("Third", new[] { 1D / 3D }, OfficeColor.Blue);
        document.AddParagraph().AddShape(OfficeLength.Points(20D), OfficeLength.Points(10D), OfficeColor.Parse("#00A060")).Rotation = 100D / 3D;
        document.CustomProperties.Add("Ratio", new WordCustomProperty(1D / 3D));

        using var package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);

        Assert.Equal("0.3333333333333333", package.MainDocumentPart!.ChartParts.Single().ChartSpace
            .Descendants<C.NumericPoint>().Single().NumericValue!.Text);
        Assert.Contains("rotation:33.333333333333336", package.MainDocumentPart.Document!.OuterXml, StringComparison.Ordinal);
        Assert.Equal("0.3333333333333333", package.CustomFilePropertiesPart!.Properties!.Descendants<VTFloat>().Single().Text);
    }
}
