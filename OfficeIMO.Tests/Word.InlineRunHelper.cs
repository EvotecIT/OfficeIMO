using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void InlineRunHelperAddsBoldAndItalicRuns() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph();
        InlineRunHelper.AddInlineRuns(paragraph, "Hello **world** and *universe*");

        var runs = paragraph._paragraph.Elements<Run>().ToList();
        Assert.Equal(4, runs.Count);
        Assert.Equal("Hello ", runs[0].InnerText);
        Assert.Null(runs[0].RunProperties?.Bold);
        Assert.Equal("world", runs[1].InnerText);
        Assert.NotNull(runs[1].RunProperties?.Bold);
        Assert.Equal(" and ", runs[2].InnerText);
        Assert.Null(runs[2].RunProperties?.Bold);
        Assert.Equal("universe", runs[3].InnerText);
        Assert.NotNull(runs[3].RunProperties?.Italic);
        Assert.Null(runs[3].RunProperties?.Bold);
    }
}
