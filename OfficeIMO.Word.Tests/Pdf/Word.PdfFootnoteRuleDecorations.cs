using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Theory]
    [InlineData("double")]
    [InlineData("dotted")]
    public void SaveAsPdf_UnqualifiedSeparatorDecorationReportsLossWhileRetainingTheRule(string pattern) {
        using WordDocument document = LoadFootnoteStorySelectionFixture();
        SetFootnoteStorySelectionMode(document, "15");
        var main = document._wordprocessingDocument.MainDocumentPart!;
        var run = (W.Run)main.FootnotesPart!.Footnotes!.Descendants<W.SeparatorMark>().Single().Parent!;
        run.RunProperties ??= new W.RunProperties();
        run.RunProperties.Underline = new W.Underline {
            Val = pattern == "double" ? W.UnderlineValues.Double : W.UnderlineValues.Dotted
        };
        Assert.Empty(document.ValidateDocument());
        string[] before = { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml,
            main.DocumentSettingsPart!.Settings!.OuterXml, main.StyleDefinitionsPart!.Styles!.OuterXml };

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "NativeFootnoteSeparatorDecorationOmitted");
        Assert.Equal("footnote separator", warning.Source);
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.DoesNotContain(result.Warnings, item => item.Code is "NativeFootnoteSeparatorMarkOmitted" or
            "NativeFootnotePlacementApproximated");
        using var pdf = PdfPigDocument.Open(actual);
        AssertFootnoteStorySelectionInventory(pdf);
        foreach (var page in pdf.GetPages()) {
            var story = PageFootnoteToken(page, "START SEPARATOR");
            Assert.Contains(page.Paths, path => path.IsFilled && path.GetBoundingRectangle() is { } band &&
                Math.Abs(band.Width - 144D) < .001D && band.Height > 0D && band.Height < 2D &&
                band.Left >= story[story.Length - 1].EndBaseLine.X &&
                band.Left < story[story.Length - 1].EndBaseLine.X + story[0].PointSize &&
                band.Bottom > story[0].StartBaseLine.Y && band.Top < story[0].BoundingBox.Top);
        }
        Assert.Equal(before, new[] { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml,
            main.DocumentSettingsPart.Settings.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml });
        Assert.Empty(document.ValidateDocument());
    }
}
