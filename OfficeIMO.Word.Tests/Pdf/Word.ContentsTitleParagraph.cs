using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordContentsTitleParagraphTests {
    [Fact]
    public void StyledTitleSurvivesRefreshAndSerializationAndStartsANewPdfPage() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Cover text");
        WordTableOfContent contents = document.AddTableOfContent();
        contents.Text = "Contents";
        WordParagraph title = Assert.IsType<WordParagraph>(contents.TitleParagraph);
        title.Style = WordParagraphStyles.Heading2;
        title.PageBreakBefore = true;
        title.FontSize = 14;
        title.ColorHex = "202428";
        WordParagraph heading = document.AddParagraph("Evidence");
        heading.Style = WordParagraphStyles.Heading1;
        heading.PageBreakBefore = true;
        contents.RefreshEntries();
        contents.RefreshEntries();
        using var serialized = new MemoryStream();
        document.Save(serialized);
        serialized.Position = 0;
        using WordDocument reopened = WordDocument.Load(serialized);
        WordTableOfContent restored = Assert.IsType<WordTableOfContent>(reopened.TableOfContent);
        Assert.Equal("Contents", restored.Text);
        Assert.Equal(WordParagraphStyles.Heading2, restored.TitleParagraph!.Style);
        Assert.True(restored.TitleParagraph.PageBreakBefore);
        Assert.Equal(14, restored.TitleParagraph.FontSize);
        Assert.Equal("202428", restored.TitleParagraph.ColorHex);
        Assert.Single(restored.RefreshEntries().Entries);
        PdfReadDocument pdf = PdfReadDocument.Open(reopened.ToPdfDocument(new WordToPdfOptions { IncludePageNumbers = false }).ToBytes());
        Assert.True(pdf.Pages.Count == 3, string.Join(" | ", pdf.Pages.Select(page => page.ExtractText())));
        Assert.Contains("Cover text", pdf.Pages[0].ExtractText());
        Assert.DoesNotContain("Contents", pdf.Pages[0].ExtractText());
        Assert.Contains("Contents", pdf.Pages[1].ExtractText());
        Assert.Contains("Evidence", pdf.Pages[2].ExtractText());
    }
}
