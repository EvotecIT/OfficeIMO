using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PaperAddMethods_ReturnEditableAttachedCoverAndRunningParagraphs() {
        using WordDocument document = WordDocument.Create();
        var paper = new WordPaperStyle();
        var cover = paper.AddCover(document, "Assessment", "Original title", "Original subtitle");
        cover.Eyebrow.Text = "UPDATED LABEL";
        cover.Title.Text = "Edited title";
        cover.Subtitle.Text = "Edited subtitle";
        var chrome = paper.AddPageChrome(document, "Report", "Product", "Scope");
        chrome.DefaultHeader.Text = "Default running title";
        chrome.EvenHeader.Text = "Even running title";
        chrome.DefaultFooter.AddText("Default footer suffix");
        chrome.EvenFooter.AddText("Even footer suffix");

        using var bytes = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal(new[] { "UPDATED LABEL", "Edited title", "Edited subtitle" }, reopened.Paragraphs.Select(p => p.Text));
        Assert.Contains(reopened.PageSetup.Header.Paragraphs, p => p.Text == "Default running title");
        Assert.Contains(reopened.PageSetup.EvenPageHeader.Paragraphs, p => p.Text == "Even running title");
        Assert.Contains(reopened.PageSetup.Footer.Paragraphs, p => p.Text.Contains("Default footer suffix"));
        Assert.Contains(reopened.PageSetup.EvenPageFooter.Paragraphs, p => p.Text.Contains("Even footer suffix"));
        Assert.True(reopened.PageSetup.DifferentOddAndEvenPages);
        Assert.Empty(reopened.ValidateDocument());
    }
}
