using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordTableStyleIdentityTests {
    [Fact]
    public void CaseDistinctNormalTableStyleRetainsItsInheritedShading() {
        using MemoryStream input = CreateCaseDistinctNormalTableDocument();
        using WordDocument document = WordDocument.Load(input);
        string before = document.MainDocumentPartRoot.Document.OuterXml;

        WordTable table = Assert.Single(document.Tables);
        Assert.Null(table.Style);
        Assert.Equal("FF0000", WordRunFormatting.ResolveTableShading(table)?.Fill?.Value);

        Assert.Equal(before, document.MainDocumentPartRoot.Document.OuterXml);
    }

    [Fact]
    public void NativePdfPaintsCaseDistinctNormalTableStyle() {
        using MemoryStream input = CreateCaseDistinctNormalTableDocument();
        using WordDocument document = WordDocument.Load(input);
        string before = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.OuterXml;

        using var model = document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(model.ToBytes(new PdfSaveOptions { AllowLoss = true }));
        var page = pdf.GetPage(1);
        Assert.Contains("CELL", page.Text, StringComparison.Ordinal);
        var firstLetter = page.Letters.First(letter => letter.Value == "C");
        Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 0D, 0D) &&
            path.GetBoundingRectangle() is { } bounds && bounds.Left <= firstLetter.StartBaseLine.X &&
            bounds.Right >= firstLetter.StartBaseLine.X && bounds.Bottom <= firstLetter.StartBaseLine.Y &&
            bounds.Top >= firstLetter.StartBaseLine.Y);
        Assert.Equal(before, document.MainDocumentPartRoot.StyleDefinitionsPart.Styles.OuterXml);
    }

    private static MemoryStream CreateCaseDistinctNormalTableDocument() {
        var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document, true)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new Document(new Body(new Table(
                new TableProperties(new TableStyle { Val = "tablenormal" }, new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "2880" }),
                new TableGrid(new GridColumn { Width = "2880" }),
                new TableRow(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2880" }),
                    new Paragraph(new Run(new Text("CELL"))))))));
            var authored = new Style(new StyleName { Val = "Authored lowercase table" }) {
                Type = StyleValues.Table, StyleId = "tablenormal", StyleTableProperties = new StyleTableProperties()
            };
            authored.StyleTableProperties.AddChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" }, true);
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Normal Table" }) { Type = StyleValues.Table, StyleId = "TableNormal", Default = true }, authored);
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(error => error.Description));
        }
        input.Position = 0;
        return input;
    }
}
