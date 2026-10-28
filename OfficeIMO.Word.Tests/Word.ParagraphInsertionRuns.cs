using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body")]
    [InlineData("section")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footer")]
    public void AppendedParagraphTextStartsWithItsAuthoredRun(string story) {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordParagraph paragraph = story switch {
            "body" => document.AddParagraph(),
            "section" => document.Sections[0].AddParagraph(),
            "cell" => document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph(removeExistingParagraphs: true),
            "header" => document.Header.Default!.AddParagraph(),
            "footer" => document.Footer.Default!.AddParagraph(),
            _ => throw new ArgumentOutOfRangeException(nameof(story))
        };
        paragraph.AddText("FIRST").SetBold();
        paragraph.AddText(" SECOND");

        using var stream = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        IEnumerable<Paragraph> paragraphs = story switch {
            "header" => package.MainDocumentPart!.HeaderParts.SelectMany(part => part.Header.Elements<Paragraph>()),
            "footer" => package.MainDocumentPart!.FooterParts.SelectMany(part => part.Footer.Elements<Paragraph>()),
            _ => package.MainDocumentPart!.Document.Descendants<Paragraph>()
        };
        Paragraph saved = Assert.Single(paragraphs, item => item.InnerText == "FIRST SECOND");
        Run[] runs = saved.Elements<Run>().ToArray();
        Assert.Equal(2, runs.Length);
        Assert.Equal("FIRST", runs[0].InnerText);
        Assert.NotNull(runs[0].RunProperties?.Bold);
        Assert.Equal(" SECOND", runs[1].InnerText);
        Assert.Null(runs[1].RunProperties?.Bold);
    }
}
