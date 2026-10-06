using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("chart", false)]
    [InlineData("chart", true)]
    [InlineData("vml", false)]
    [InlineData("vml", true)]
    public void SaveAsPdf_MixedPictureRunPreservesFollowingDrawable(string kind, bool leadingText) {
        string sourcePath = Path.Combine(_directoryWithFiles, $"PdfMixedRunDrawable-{kind}-{leadingText}.docx");
        using WordDocument document = WordDocument.Create(sourcePath);
        WordParagraph paragraph = document.AddParagraph(leadingText ? "BeforeDrawable" : string.Empty);
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "inline.png", 16, 32);
        W.Run sourceRun = paragraph._run!;
        W.Run drawableRun;
        if (kind == "chart") {
            WordChart chart = document.AddChart("MixedRunChart", false, 240, 160);
            chart.AddPie("One", 3);
            chart.AddPie("Two", 4);
            drawableRun = chart.Drawing!.Ancestors<W.Run>().Single();
        } else {
            WordShape shape = document.AddShape(WordShapeType.Rectangle, 36, 18, "#CCB399", "#1A334D", 2.5);
            drawableRun = shape._run;
        }
        W.Paragraph drawableParagraph = drawableRun.Ancestors<W.Paragraph>().Single();
        foreach (var child in drawableRun.ChildElements.Where(child => child is not W.RunProperties).ToArray()) {
            child.Remove();
            sourceRun.Append(child);
        }
        drawableParagraph.Remove();
        sourceRun.Append(new W.Text("AfterDrawable"));
        document.Save();
        string sourceXml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Equal(sourceXml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Single(pdf.GetPages().SelectMany(page => page.GetImages()));
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("AfterDrawable", text);
        if (leadingText) Assert.Contains("BeforeDrawable", text);
        if (kind == "chart") Assert.Single(words, word => word.Text == "MixedRunChart");
        else Assert.Single(System.Text.RegularExpressions.Regex.Matches(PdfOperatorSearchText.From(bytes), "0\\.8 0\\.702 0\\.6 rg").Cast<System.Text.RegularExpressions.Match>());
    }
}
