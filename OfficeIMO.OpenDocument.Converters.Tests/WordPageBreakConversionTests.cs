using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordPageBreakConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LaterImageMetadataLossIsReportedAndRejectedInStrictMode(bool header) {
        using var source = WordDocument.Create();
        if (header) source.AddHeadersAndFooters();
        var paragraph = header ? source.Sections[0].Header.Default!.AddParagraph() : source.AddParagraph();
        byte[] bytes = OfficeIMO.Drawing.OfficePngWriter.Encode(new OfficeIMO.Drawing.OfficeRasterImage(2, 2, OfficeIMO.Drawing.OfficeColor.Red));
        using (var first = new System.IO.MemoryStream(bytes)) paragraph.AddImage(first, "first.png", 12, 12);
        using (var second = new System.IO.MemoryStream(bytes)) paragraph.AddImage(second, "second.png", 12, 12);
        paragraph.GetPositionedImages().Last().Image.Description = "Second image description";
        var result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "image-layout" && mapping.Count == 1);
        Assert.Throws<OfficeIMO.OpenDocument.OdfConversionLossException>(() => source.ToOpenDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OfficeIMO.OpenDocument.OdfConversionLossPolicy.ThrowOnAnyLoss
        }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineImageKeepsItsPositionRelativeToAPageBreakInTheSameRun(bool breakBeforeImage) {
        using var source = WordDocument.Create();
        var paragraph = source.AddParagraph();
        var raster = new OfficeIMO.Drawing.OfficeRasterImage(2, 2, OfficeIMO.Drawing.OfficeColor.Red);
        using var image = new System.IO.MemoryStream(OfficeIMO.Drawing.OfficePngWriter.Encode(raster));
        paragraph.AddImage(image, "marker.png", 12, 12);
        var drawing = paragraph._run!.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>()!;
        drawing.InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Text("BEFORE"));
        var boundary = new DocumentFormat.OpenXml.Wordprocessing.Break { Type = DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page };
        if (breakBeforeImage) drawing.InsertBeforeSelf(boundary);
        else drawing.InsertAfterSelf(boundary);
        paragraph._run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text("AFTER"));
        var result = source.ToOpenDocumentResult();
        var paragraphs = result.Value.Paragraphs;
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("BEFORE", paragraphs[0].Text);
        Assert.Equal("AFTER", paragraphs[1].Text);
        Assert.Single(paragraphs[breakBeforeImage ? 1 : 0].Images);
        Assert.Empty(paragraphs[breakBeforeImage ? 0 : 1].Images);
        Assert.True(paragraphs[1].PageBreakBefore);
    }

    [Fact]
    public void MultipleImagesAndTextRetainSourceOrderAcrossTheBreak() {
        using var source = WordDocument.Create();
        var paragraph = source.AddParagraph("A");
        byte[] bytes = OfficeIMO.Drawing.OfficePngWriter.Encode(new OfficeIMO.Drawing.OfficeRasterImage(2, 2, OfficeIMO.Drawing.OfficeColor.Red));
        using (var first = new System.IO.MemoryStream(bytes)) paragraph.AddImage(first, "first.png", 12, 12);
        paragraph._run!.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text("B"));
        using (var second = new System.IO.MemoryStream(bytes)) paragraph.AddImage(second, "second.png", 12, 12);
        paragraph._run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text("C"));
        paragraph._run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Break { Type = DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page });
        using (var third = new System.IO.MemoryStream(bytes)) paragraph.AddImage(third, "third.png", 12, 12);
        paragraph._run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text("D"));
        var paragraphs = source.ToOpenDocumentResult().Value.Paragraphs;
        Assert.Equal(2, paragraphs.Count);
        string Order(OdtParagraph item) => string.Concat(item.InlineNodes.Select(node => node.Image != null ? "[IMAGE]" : node.Text));
        Assert.Equal("A[IMAGE]B[IMAGE]C", Order(paragraphs[0]));
        Assert.Equal("[IMAGE]D", Order(paragraphs[1]));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordPageBreaksSurviveTheOdtRoundTrip(bool insideParagraph) {
        using var source = WordDocument.Create();
        var first = source.AddParagraph("Before the break");
        if (insideParagraph) first.AddBreak(WordBreakType.Page).AddText("After the break");
        else { source.AddPageBreak(); source.AddParagraph("After the break"); }
        var odt = source.ToOpenDocumentResult().Value;
        using var converted = odt.ToWordDocumentResult().Value;
        Assert.Equal(2, converted.GetEstimatedPageCount());
        Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text.Contains("Before the break"));
        Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text.Contains("After the break"));
    }

    [Fact]
    public void ConsecutivePageBreaksPreserveTheBlankPageAndRunFormatting() {
        using var source = WordDocument.Create();
        var paragraph = source.AddParagraph("Before");
        paragraph.Bold = true;
        paragraph.AddBreak(WordBreakType.Page).AddBreak(WordBreakType.Page).AddText("After").Bold = true;
        using var converted = source.ToOpenDocumentResult().Value.ToWordDocumentResult().Value;
        Assert.Equal(3, converted.GetEstimatedPageCount());
        Assert.Contains(converted.Paragraphs, item => item.Text == "Before" && item.Bold == true);
        Assert.Contains(converted.Paragraphs, item => item.Text == "After" && item.Bold == true);
    }

    [Fact]
    public void LiteralLineSeparatorTextDoesNotCreateAPageBreak() {
        using var source = WordDocument.Create();
        source.AddParagraph("Before\u2028After");
        using var converted = source.ToOpenDocumentResult().Value.ToWordDocumentResult().Value;
        Assert.Equal(1, converted.GetEstimatedPageCount());
        Assert.Contains(converted.Paragraphs, item => item.Text.Contains("Before") && item.Text.Contains("After"));
    }
}
