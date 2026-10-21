using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfBiDiOnOffTests {
    [Theory]
    [InlineData("missing", false)]
    [InlineData("false", false)]
    [InlineData("0", false)]
    [InlineData("true", true)]
    [InlineData("implicit", true)]
    public void SaveAsPdf_BiDiOnOffControlsDefaultAlignmentAndViewerDirection(string sourceValue, bool rightToLeft) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("TARGET");
        W.Paragraph paragraph = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().Last();
        SetSourceBiDi(paragraph, sourceValue);

        byte[] bytes = ConvertBiDiDocumentWithoutSourceChanges(document);
        using var pdf = PdfPigDocument.Open(bytes);
        var target = Assert.Single(Assert.Single(pdf.GetPages()).GetWords(), word => word.Text == "TARGET");
        if (rightToLeft)
            Assert.InRange(target.Letters.Last().EndBaseLine.X, 269.99D, 270.01D);
        else
            Assert.InRange(target.Letters.First().StartBaseLine.X, 29.99D, 30.01D);
        Assert.Equal(rightToLeft ? "R2L" : null, PdfInspector.Inspect(bytes).ViewerPreferences?.GetValue("Direction"));
    }

    [Theory]
    [InlineData("false", false)]
    [InlineData("true", true)]
    public void SaveAsPdf_BiDiRetainsExplicitAlignmentAndConfiguredViewerDirection(string sourceValue, bool explicitLeft) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("EXPLICIT");
        W.Paragraph paragraph = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().Last();
        SetSourceBiDi(paragraph, sourceValue);
        paragraph.ParagraphProperties!.Justification = new W.Justification {
            Val = explicitLeft ? W.JustificationValues.Left : W.JustificationValues.Right
        };

        byte[] bytes = ConvertBiDiDocumentWithoutSourceChanges(document, PdfViewerDirection.LeftToRight);
        using var pdf = PdfPigDocument.Open(bytes);
        var target = Assert.Single(Assert.Single(pdf.GetPages()).GetWords(), word => word.Text == "EXPLICIT");
        if (explicitLeft)
            Assert.InRange(target.Letters.First().StartBaseLine.X, 29.99D, 30.01D);
        else
            Assert.InRange(target.Letters.Last().EndBaseLine.X, 269.99D, 270.01D);
        Assert.Equal("L2R", PdfInspector.Inspect(bytes).ViewerPreferences?.GetValue("Direction"));
    }

    [Theory]
    [InlineData("table")]
    [InlineData("footnote")]
    public void SaveAsPdf_BiDiFalseKeepsContainerParagraphsAtTheControlLeftEdge(string container) {
        using WordDocument document = WordDocument.Create();
        W.Paragraph control = new W.Paragraph(new W.Run(new W.Text("CONTROL")));
        W.Paragraph target = new W.Paragraph(new W.Run(new W.Text("TARGET")));
        SetSourceBiDi(target, "0");
        var main = document.OpenXmlDocument.MainDocumentPart!;
        if (container == "table") {
            document.AddTable(1, 1);
            W.TableCell cell = main.Document.Body!.Descendants<W.TableCell>().Single();
            cell.RemoveAllChildren<W.Paragraph>();
            cell.Append(control, target);
        } else {
            document.AddParagraph("REFERENCE").AddFootNote("FIRSTNOTE");
            W.Footnote note = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()
                .Single(item => item.Type == null && item.Id?.Value > 0);
            note.Append(control, target);
        }

        byte[] bytes = ConvertBiDiDocumentWithoutSourceChanges(document);
        using var pdf = PdfPigDocument.Open(bytes);
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var expected = Assert.Single(words, word => word.Text == "CONTROL");
        var actual = Assert.Single(words, word => word.Text == "TARGET");
        Assert.InRange(actual.Letters.First().StartBaseLine.X,
            expected.Letters.First().StartBaseLine.X - .01D, expected.Letters.First().StartBaseLine.X + .01D);
        Assert.Null(PdfInspector.Inspect(bytes).ViewerPreferences?.GetValue("Direction"));
    }

    private static void SetSourceBiDi(W.Paragraph paragraph, string sourceValue) {
        if (sourceValue == "missing") return;
        W.BiDi bidi = new W.BiDi();
        if (sourceValue != "implicit")
            bidi.SetAttribute(new OpenXmlAttribute("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", sourceValue));
        paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        paragraph.ParagraphProperties.BiDi = bidi;
    }

    private static byte[] ConvertBiDiDocumentWithoutSourceChanges(WordDocument document, PdfViewerDirection? viewerDirection = null) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        Assert.Empty(document.ValidateDocument());
        string bodyBefore = main.Document.OuterXml;
        string? notesBefore = main.FootnotesPart?.Footnotes?.OuterXml;
        string? stylesBefore = main.StyleDefinitionsPart?.Styles?.OuterXml;
        string? settingsBefore = main.DocumentSettingsPart?.Settings?.OuterXml;
        var options = new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new PageSize(300, 300),
            Margins = PageMargins.Uniform(30),
            FontFamily = "Helvetica"
        };
        if (viewerDirection.HasValue)
            options.PdfOptions = new PdfOptions {
                ViewerPreferences = new PdfViewerPreferencesOptions { Direction = viewerDirection.Value }
            };
        byte[] bytes = document.ToPdfDocumentResult(options).ToBytes();
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Equal(bodyBefore, main.Document.OuterXml);
        Assert.Equal(notesBefore, main.FootnotesPart?.Footnotes?.OuterXml);
        Assert.Equal(stylesBefore, main.StyleDefinitionsPart?.Styles?.OuterXml);
        Assert.Equal(settingsBefore, main.DocumentSettingsPart?.Settings?.OuterXml);
        Assert.Empty(document.ValidateDocument());
        return bytes;
    }
}
