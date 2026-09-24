using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordBreakOnlyTextTests {
    [Theory]
    [InlineData("\n", 1)]
    [InlineData("\r\n", 1)]
    [InlineData("\n\n", 2)]
    public void AddFormattedTextReturnsAParagraphForBreakOnlyInput(string text, int expectedBreaks) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Before");

        WordParagraph added = paragraph.AddFormattedText(text, bold: true);
        added.SetLanguage("en");

        Assert.True(added.IsBreak);
        Run[] breakRuns = document._wordprocessingDocument.MainDocumentPart!.Document
            .Descendants<Run>().Where(run => run.Elements<Break>().Any()).ToArray();
        Assert.Equal(expectedBreaks, breakRuns.Length);
        Assert.All(breakRuns, run => {
            Assert.Single(run.Elements<Break>());
            Assert.NotNull(run.RunProperties?.Bold);
        });
        Assert.Equal("en", breakRuns[^1].RunProperties?.Languages?.Val?.Value);
    }
}
