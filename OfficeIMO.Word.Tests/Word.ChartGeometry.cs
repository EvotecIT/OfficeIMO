using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordChartLengthTests {
    [Fact]
    public void ChartFramesRetainExactNativeDimensionsThroughCreationResizeAndReload() {
        using var document = WordDocument.Create();
        var chart = document.AddChart("Exact geometry", width: OfficeLength.Emus(1234567), height: OfficeLength.Emus(2345678));
        chart.AddCategories(new() { "A", "B" });
        chart.AddBar("Values", new[] { 1, 2 }, OfficeColor.Blue);
        Assert.Equal(1234567, chart.Drawing!.Descendants<DW.Extent>().Single().Cx!.Value);
        chart.SetSize(OfficeLength.Emus(3456789));
        byte[] bytes = document.ToBytes();
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var extent = package.MainDocumentPart!.Document.Descendants<DW.Extent>().Single();
        Assert.Equal(3456789, extent.Cx!.Value);
        Assert.Equal(2345678, extent.Cy!.Value);
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(package).ToArray();
        Assert.True(errors.Length == 0, string.Join(System.Environment.NewLine, errors.Select(error => error.Description)));
        using var loaded = WordDocument.Load(new MemoryStream(bytes));
        Assert.Single(loaded.Charts).SetSize(OfficeLength.Centimeters(2), OfficeLength.Points(30));
        var loadedExtent = Assert.Single(loaded.Charts).Drawing!.Descendants<DW.Extent>().Single();
        Assert.Equal(720000, loadedExtent.Cx!.Value);
        Assert.Equal(381000, loadedExtent.Cy!.Value);
    }
}
