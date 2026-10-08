using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FindAndReplace_UsesSignificantXmlTextForContainedAndCrossRunMatches(bool crossRun) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph(crossRun ? "First " : " First ");
        if (crossRun) {
            paragraph._paragraph.RemoveAllChildren<W.Run>();
            paragraph._paragraph.Append(new W.Run(new W.Text("First ")), new W.Run(new W.Text(" Second")));
        }
        foreach (W.Text text in paragraph._paragraph.Descendants<W.Text>()) text.Space = null;
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string target = crossRun ? "FirstSecond" : "First";
        Assert.Equal(target, string.Concat(imported.Paragraphs.Select(value => value.Text)));
        Assert.NotEmpty(imported.Find(target, StringComparison.Ordinal));
        Assert.Equal(1, imported.FindAndReplace(target, "Updated", StringComparison.Ordinal));
        Assert.Equal("Updated", string.Concat(imported.Paragraphs.Select(value => value.Text)));
    }

    [Theory]
    [InlineData(" \t", false, false, false)]
    [InlineData(" \t", true, false, false)]
    [InlineData("Before \t", false, false, false)]
    [InlineData("Before \t", true, false, false)]
    [InlineData(" \t", false, true, false)]
    [InlineData(" \t", true, true, false)]
    [InlineData("Before \t", false, true, false)]
    [InlineData("Before \t", true, true, false)]
    [InlineData(" \t", false, false, true)]
    [InlineData(" \t", true, false, true)]
    [InlineData("Before \t", false, false, true)]
    [InlineData("Before \t", true, false, true)]
    [InlineData(" \t", false, true, true)]
    [InlineData(" \t", true, true, true)]
    [InlineData("Before \t", false, true, true)]
    [InlineData("Before \t", true, true, true)]
    public void ParagraphText_IgnoredXmlEdgesDoNotMovePageOrColumnBreaks(string prefix, bool column, bool replace, bool wrapping) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("seed");
        W.Run run = paragraph._run!;
        run.RemoveAllChildren();
        W.BreakValues breakType = column ? W.BreakValues.Column : W.BreakValues.Page;
        run.Append(new W.Text(prefix), new W.Break { Type = breakType }, new W.Text(" After "));
        if (wrapping) run.InsertBefore(new W.Break(), run.Elements<W.Break>().Single());
        string before = prefix.Trim(' ', '\t') + (wrapping ? "\n" : "") + "\u2028After";
        Assert.Equal(before, paragraph.Text);
        if (replace) Assert.Equal(1, document.FindAndReplace("After", "Updated", StringComparison.Ordinal));
        else paragraph.Text = paragraph.Text;
        Assert.Equal(replace ? before.Replace("After", "Updated") : before, paragraph.Text);
        Assert.Equal(breakType, Assert.Single(run.Elements<W.Break>(), value => value.Type?.Value == breakType).Type!.Value);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("header")]
    public void RepeatingSection_ImportedWhitespaceMatchesPublicAndPdfText(string story) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (story == "header") {
            document.AddHeadersAndFooters();
            paragraph = RequireSectionHeader(document, 0, W.HeaderFooterValues.Default).AddParagraph();
        } else if (story == "table") paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        else paragraph = document.AddParagraph();
        WordRepeatingSection section = paragraph.AddRepeatingSection("Seed", "Whitespace");
        section.SetTextItems(new[] { "Seed" });
        W.Run run = section._sdtRun.Descendants<W.Run>().Single();
        run.RemoveAllChildren();
        run.Append(new W.Text(" First "), new W.Text(" Second "));
        string xml = section._sdtRun.OuterXml;
        Assert.Equal("FirstSecond", Assert.Single(section.TextItems));
        Assert.Equal(xml, section._sdtRun.OuterXml);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal("FirstSecond", string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value)));
        Assert.Equal(xml, section._sdtRun.OuterXml);
    }

    [Fact]
    public void RepeatingSection_UnknownXmlTextHonorsWhitespaceAndPreservedReplacement() {
        using WordDocument document = WordDocument.Create();
        WordRepeatingSection section = document.AddParagraph().AddRepeatingSection("Seed", "Whitespace");
        OpenXmlElement item = section._sdtRun.SdtContentRun!.FirstChild!;
        item.RemoveAllChildren();
        item.InnerXml = "<w:sdt xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:sdtContent><w:r><w:t> First </w:t><w:t> Second </w:t></w:r></w:sdtContent></w:sdt>";
        Assert.Empty(item.Descendants<W.Text>());
        string xml = section._sdtRun.OuterXml;
        Assert.Equal("FirstSecond", Assert.Single(section.TextItems));
        Assert.Equal(xml, section._sdtRun.OuterXml);
        item.InnerXml = item.InnerXml.Replace("<w:t>", "<w:t xml:space=\"preserve\">");
        Assert.Equal(" First  Second ", Assert.Single(section.TextItems));
        section.SetTextItems(new[] { " Preserved " });
        Assert.Equal(" Preserved ", Assert.Single(section.TextItems));
    }

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
