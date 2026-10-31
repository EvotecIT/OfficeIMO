using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordCharacterSpacingPhysicalLengthTests {
    [Theory]
    [InlineData(318L, 1)]
    [InlineData(-318L, -1)]
    [InlineData(0L, 0)]
    [InlineData(1363652115845L, int.MaxValue)]
    [InlineData(-1363652116480L, int.MinValue)]
    public void CharacterSpacingWritesSignedTwipsAndCanBeCleared(long emus, int expectedTwips) {
        using var document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Physical tracking");
        string initial = run._paragraph.OuterXml;
        Assert.Null(run.Spacing);
        Assert.Equal(initial, run._paragraph.OuterXml);
        run.SetSpacing(OfficeLength.Emus(emus));
        Assert.Equal(expectedTwips, run._run!.RunProperties!.Spacing!.Val!.Value);
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using var loaded = WordDocument.Load(stream);
        WordParagraph actual = loaded.Paragraphs[0];
        Assert.Equal(OfficeLength.Emus((long)expectedTwips * 635), actual.Spacing);
        actual.Spacing = null;
        Assert.Null(actual._run!.RunProperties!.GetFirstChild<Spacing>());
    }

    [Fact]
    public void OutOfRangeCharacterSpacingDoesNotChangeTheRun() {
        using var document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Keep tracking");
        run.Spacing = OfficeLength.Points(-0.5);
        string before = run._paragraph.OuterXml;
        Assert.Throws<OverflowException>(() => run.Spacing = OfficeLength.Emus(((long)int.MaxValue + 1) * 635));
        Assert.Equal(before, run._paragraph.OuterXml);
        Assert.Throws<OverflowException>(() => run.SetSpacing(OfficeLength.Emus(((long)int.MinValue - 1) * 635)));
        Assert.Equal(before, run._paragraph.OuterXml);
    }
}
