using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_AutoHeightSingleParagraphCellPreservesMeasuredWideText(bool useShaper) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        var table = document.AddTable(1, 1);
        table.Width = 2000;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new System.Collections.Generic.List<int> { 2000 };
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        var paragraph = table.Rows[0].Cells[0].AddParagraph(CreateFrameTokens(), removeExistingParagraphs: true);
        paragraph.SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        paragraph.FontSizePoints = 11D;

        var options = CreateWideTextOptions(useShaper);
        var image = document.ExportImage(OfficeImageExportFormat.Svg, options);
        WriteCellFontEvidence(document.CreateVisualSnapshot(options), image, useShaper);

        AssertPaintedTokens(new[] { Encoding.UTF8.GetString(image.Bytes) }, "F", 20);
        Assert.DoesNotContain(image.Diagnostics, item => item.Code == "unsupported-word-table-cell-text-overflow");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_VmlFitShapePreservesMeasuredWideText(bool useShaper) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.PageSize = WordPageSize.A4;
        document.Margins.Type = WordMargin.Narrow;
        var textBox = document.AddParagraph().AddTextBoxVml(CreateFrameTokens());
        foreach (var paragraph in textBox.Paragraphs) {
            paragraph.SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
            paragraph.FontSizePoints = 11D;
        }
        V.Shape shape = document.BodyRoot.Descendants<V.Shape>().Last(item => item.Descendants<V.TextBox>().Any());
        shape.Style = "width:96pt;mso-wrap-style:square;mso-fit-shape-to-text:t";

        var image = document.ExportImage(OfficeImageExportFormat.Svg, CreateWideTextOptions(useShaper));

        AssertPaintedTokens(new[] { Encoding.UTF8.GetString(image.Bytes) }, "F", 20);
    }

    private static string CreateFrameTokens() => string.Join(" ", Enumerable.Range(1, 20)
        .Select(index => "F" + index.ToString("00", CultureInfo.InvariantCulture)));
}
