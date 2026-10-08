using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    public void ParagraphText_UsesDeclaredXmlWhitespaceWithoutChangingSource(int mode, bool formatted) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("seed");
        W.Run run = paragraph._run!;
        run.RemoveAllChildren();
        if (formatted) run.Append(new W.RunProperties(new W.Bold()));
        const string value = " \t\r\nFirst  middle\u00a0last\u2003 \t";
        var text = new W.Text(value);
        if (mode == 1) text.Space = SpaceProcessingModeValues.Default;
        if (mode == 2) text.Space = SpaceProcessingModeValues.Preserve;
        run.Append(text);
        string xml = paragraph._paragraph.OuterXml;
        Assert.Equal(mode == 2 ? value : "First  middle\u00a0last\u2003", paragraph.Text);
        Assert.Equal(xml, paragraph._paragraph.OuterXml);
    }

    [Fact]
    public void HyperlinkText_ImportedWhitespaceAndAuthoredReplacementUseTheirRespectiveModes() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph().AddHyperLink(" Old ", new Uri("https://example.test/link"));
        W.Text text = paragraph._paragraph.Descendants<W.Text>().Single();
        text.Space = null;
        string xml = paragraph._paragraph.OuterXml;
        Assert.Equal("Old", paragraph.Hyperlink!.Text);
        Assert.Equal(xml, paragraph._paragraph.OuterXml);
        paragraph.Hyperlink.Text = " New ";
        Assert.Equal(" New ", paragraph.Hyperlink.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, text.Space!.Value);
    }

    [Fact]
    public void ParagraphText_XmlWhitespaceDoesNotRemoveStructuralTabsOrBreaks() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("seed");
        W.Run run = paragraph._run!;
        run.RemoveAllChildren();
        run.Append(new W.Text(" First "), new W.TabChar(), new W.Text(" Second "),
            new W.Break(), new W.Text(" Third "));
        string xml = paragraph._paragraph.OuterXml;
        Assert.Equal("First\tSecond\nThird", paragraph.Text);
        Assert.Equal(xml, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("header")]
    public void SaveAsPdf_ImportedUnpreservedRunEdgesDoNotIntroduceTextGaps(string story) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (story == "header") {
            document.AddHeadersAndFooters();
            paragraph = RequireSectionHeader(document, 0, W.HeaderFooterValues.Default).AddParagraph();
        } else if (story == "table") {
            paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        } else paragraph = document.AddParagraph();
        paragraph.Text = " First ";
        paragraph.AddText(" Second ");
        foreach (W.Text text in paragraph._paragraph.Descendants<W.Text>()) text.Space = null;
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal("FirstSecond", string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value)));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }
}
