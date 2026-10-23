using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("table", false)]
    [InlineData("footnote", false)]
    [InlineData("endnote", false)]
    [InlineData("header", false)]
    [InlineData("footer", false)]
    public void SaveAsPdf_ParagraphFramesReportApproximationAndPreserveSource(string story, bool inherited) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODY");
        WordParagraph paragraph = AddFrameDiagnosticParagraph(document, story);
        paragraph.Text = "FRAMEDMARKER";
        var frame = new W.FrameProperties();
        frame.SetAttribute(new OpenXmlAttribute("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "3600"));
        frame.SetAttribute(new OpenXmlAttribute("w", "x", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "1440"));
        frame.SetAttribute(new OpenXmlAttribute("w", "y", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "720"));
        if (inherited) {
            document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleParagraphProperties(frame)) {
                    Type = W.StyleValues.Paragraph, StyleId = "PositionedFrame"
                });
            paragraph.SetStyleId("PositionedFrame");
        } else {
            paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.AddChild(frame, true);
        }
        Assert.Empty(document.ValidateDocument());
        string paragraphXml = paragraph._paragraph.OuterXml;
        string stylesXml = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;
        var options = new WordToPdfOptions();

        for (int index = 0; index < 2; index++) {
            var result = document.ToPdfDocumentResult(options);
            PdfConversionWarning warning = Assert.Single(result.Warnings,
                item => item.Code == "NativeParagraphFrameFlattened");
            Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
            Assert.True(result.HasLoss);
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
            using var output = new MemoryStream();
            var saved = result.Save(output);
            Assert.True(saved.Succeeded);
            Assert.True(saved.HasLoss);
            Assert.Single(saved.Warnings, item => item.Code == "NativeParagraphFrameFlattened");
            byte[] savedPdf = output.ToArray();
            Assert.Contains("FRAMEDMARKER", PdfReadDocument.Open(savedPdf,
                new PdfLoadOptions { IncludeArtifactText = true }).ExtractText(), StringComparison.Ordinal);
            if (story is "header" or "footer")
                Assert.DoesNotContain("FRAMEDMARKER", PdfReadDocument.Open(savedPdf).ExtractText(), StringComparison.Ordinal);
            Assert.Equal(paragraphXml, paragraph._paragraph.OuterXml);
            Assert.Equal(stylesXml, document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml);
        }
    }

    [Fact]
    public void SaveAsPdf_OrdinaryParagraphsDoNotReportFrameLoss() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("PLAINMARKER");
        var result = document.ToPdfDocumentResult();
        Assert.DoesNotContain(result.Warnings, item => item.Code == "NativeParagraphFrameFlattened");
        Assert.Contains("PLAINMARKER", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    private static WordParagraph AddFrameDiagnosticParagraph(WordDocument document, string story) {
        switch (story) {
            case "body": return document.AddParagraph();
            case "table": return document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            case "header": return document.HeaderDefaultOrCreate.AddParagraph();
            case "footer": return document.FooterDefaultOrCreate.AddParagraph();
            default:
                WordParagraph body = document.AddParagraph("REFERENCE");
                bool endnote = story == "endnote";
                if (endnote) body.AddEndNote("NOTE");
                else body.AddFootNote("NOTE");
                var main = document.OpenXmlDocument.MainDocumentPart!;
                W.FootnoteEndnoteType note = endnote
                    ? Assert.Single(main.EndnotesPart!.Endnotes!.Elements<W.Endnote>(), item => item.Id?.Value ==
                        Assert.Single(body._paragraph.Descendants<W.EndnoteReference>()).Id!.Value)
                    : Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), item => item.Id?.Value ==
                        Assert.Single(body._paragraph.Descendants<W.FootnoteReference>()).Id!.Value);
                return new WordParagraph(document, Assert.Single(note.Elements<W.Paragraph>()));
        }
    }
}
