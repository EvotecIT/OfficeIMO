using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordBreakRunOwnershipTests {
    [Theory]
    [InlineData(null)]
    [InlineData(WordBreakType.Page)]
    [InlineData(WordBreakType.Column)]
    public void ReturnedBreakHandleFormatsTheSavedRun(WordBreakType? type) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Before");
        WordParagraph handle = paragraph.AddBreak(type);
        handle.SetLanguage("en").SetBold(true);
        Assert.True(handle.IsBreak);
        using var bytes = new MemoryStream();
        document.Save(bytes);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        WordParagraph saved = Assert.Single(reopened.Paragraphs.Where(item => item.IsBreak));
        Assert.Equal("en", saved.Language);
        Assert.True(saved.Bold);
        Assert.Contains("Before", reopened.Paragraphs.Select(item => item.Text));
    }
}
