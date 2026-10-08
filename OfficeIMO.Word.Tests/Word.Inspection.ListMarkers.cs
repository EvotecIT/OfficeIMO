using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void InspectionListMarkersRejectOverflowBeforePublishingAnIndex() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(int.MaxValue));
        list.AddItem("Last index");
        WordParagraphSnapshot paragraph = Assert.Single(Assert.Single(document.CreateInspectionSnapshot().Sections).Elements.OfType<WordParagraphSnapshot>());
        Assert.Equal(int.MaxValue, paragraph.ListIndex);
        Assert.Equal("2147483647.", paragraph.ListMarker);
        list.AddItem("Overflow");
        Assert.Empty(document.ValidateDocument());
        Assert.Throws<System.IO.InvalidDataException>(() => document.CreateInspectionSnapshot());
    }

    [Fact]
    public void InspectionListMarkersRetainNestedParentNumbersAndLevelIndices() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.LowerLetterDot));
        list.Numbering.Levels[0].StartNumberingValue = 12;
        list.Numbering.Levels[1].StartNumberingValue = 3;
        list.Numbering.Levels[1].LevelText = "%1.%2)";
        list.AddItem("Parent");
        list.AddItem("Child", 1);
        list.AddItem("Next parent");

        WordParagraphSnapshot[] paragraphs = Assert.Single(document.CreateInspectionSnapshot().Sections).Elements.OfType<WordParagraphSnapshot>().ToArray();
        Assert.Equal(new[] { "12.", "12.c)", "13." }, paragraphs.Select(paragraph => paragraph.ListMarker));
        Assert.Equal(new int?[] { 12, 3, 13 }, paragraphs.Select(paragraph => paragraph.ListIndex));
        Assert.Equal(new int?[] { 0, 1, 0 }, paragraphs.Select(paragraph => paragraph.ListLevel));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InspectionListMarkersPreserveSharedCountersAndExplicitRestarts(bool restart) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second);
        if (restart) second.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 3 }) { LevelIndex = 0 });
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml + numbering.OuterXml;

        WordParagraphSnapshot[] snapshots = document.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements.OfType<WordParagraphSnapshot>()).ToArray();

        int[] expected = restart ? new[] { 12, 13, 3, 4 } : new[] { 12, 13, 14, 15 };
        Assert.Equal(expected.Select(value => (int?)value), snapshots.Select(paragraph => paragraph.ListIndex));
        Assert.Equal(expected.Select(value => value + "."), snapshots.Select(paragraph => paragraph.ListMarker));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml + numbering.OuterXml);
        numbering.Elements<AbstractNum>().Single().Elements<Level>().First().StartNumberingValue!.Val = 50;
        Assert.Equal(expected.Select(value => value + "."), snapshots.Select(paragraph => paragraph.ListMarker));
    }

    [Fact]
    public void InspectionListMarkersFollowBodyTableOrderAndKeepHeaderCountersSeparate() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.Levels[0].StartNumberingValue = 12;
        Attach(document.AddParagraph("Body first"));
        Attach(document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]);
        Attach(document.AddParagraph("Body last"));
        Attach(document.Header.Default.AddParagraph("Header first"));
        Attach(document.Header.Default.AddParagraph("Header second"));

        WordSectionSnapshot section = Assert.Single(document.CreateInspectionSnapshot().Sections);
        Assert.Equal(new[] { "12.", "14." }, section.Elements.OfType<WordParagraphSnapshot>().Select(paragraph => paragraph.ListMarker));
        WordTableSnapshot table = Assert.Single(section.Elements.OfType<WordTableSnapshot>());
        Assert.Equal("13.", Assert.Single(table.Rows[0].Cells[0].Paragraphs).ListMarker);
        Assert.Equal(new[] { "12.", "13." }, section.DefaultHeader!.Elements.OfType<WordParagraphSnapshot>().Where(paragraph => paragraph.IsListItem).Select(paragraph => paragraph.ListMarker));
        Assert.Empty(document.ValidateDocument());

        void Attach(WordParagraph paragraph) => paragraph._paragraph.ParagraphProperties = new ParagraphProperties(
            new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId }));
    }

    [Theory]
    [InlineData(WordListLevelKind.BulletSolidRound, "•")]
    [InlineData(WordListLevelKind.None, "")]
    public void InspectionListMarkersDistinguishBulletsInvisibleMarkersAndPlainParagraphs(WordListLevelKind kind, string marker) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(kind));
        list.AddItem("List item");
        document.AddParagraph("Plain paragraph");

        WordParagraphSnapshot[] paragraphs = Assert.Single(document.CreateInspectionSnapshot().Sections).Elements.OfType<WordParagraphSnapshot>().ToArray();
        Assert.Equal(marker, paragraphs[0].ListMarker);
        Assert.Null(paragraphs[0].ListIndex);
        Assert.Null(paragraphs[1].ListMarker);
        Assert.Null(paragraphs[1].ListIndex);
    }
}
