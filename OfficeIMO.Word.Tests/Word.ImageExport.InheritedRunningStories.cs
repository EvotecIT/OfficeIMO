using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "default", false)]
    [InlineData(false, "first", false)]
    [InlineData(false, "even", false)]
    [InlineData(true, "default", false)]
    [InlineData(true, "first", false)]
    [InlineData(true, "even", false)]
    [InlineData(false, "default", true)]
    [InlineData(false, "first", true)]
    [InlineData(false, "even", true)]
    [InlineData(true, "default", true)]
    [InlineData(true, "first", true)]
    [InlineData(true, "even", true)]
    public void ImageExport_InheritedRunningStoriesPreserveRolesAndExplicitBlank(bool footer, string kind, bool blank) {
        using WordDocument document = CreateInheritedRunningStoryControl(footer, kind, blank ? "blank" : "next");
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        IReadOnlyList<WordDocumentVisualSnapshot> batch = document.CreateVisualSnapshots();
        Assert.Equal(3, batch.Count);
        for (int index = 0; index < batch.Count; index++) {
            bool marker = !blank || index == 0;
            marker &= kind == "default" || (kind == "first" ? index <= 1 : index == 1);
            WordDocumentVisualSnapshot single = document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = index });
            foreach (WordDocumentVisualSnapshot snapshot in new[] { batch[index], single })
                Assert.Equal(marker, snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Any(text => text.Text.Contains("MARK")));
        }
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ImageExport_MissingFirstOrEvenRoleDoesNotUseDefaultStory(bool footer, bool even) {
        using WordDocument source = WordDocument.Create();
        WordHeaderFooter story = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        story.AddParagraph("DEFAULTMARK");
        source.Sections[0].DifferentFirstPage = !even;
        if (even) source.Sections[0].DifferentOddAndEvenPages = true;
        foreach (var reference in source.Sections[0]._sectionProperties.ChildElements.Where(element =>
            element is W.HeaderReference header && header.Type?.Value != W.HeaderFooterValues.Default ||
            element is W.FooterReference bottom && bottom.Type?.Value != W.HeaderFooterValues.Default).ToArray()) reference.Remove();
        source.AddParagraph("ONE").AddBreak(WordBreakType.Page);
        source.AddParagraph("TWO");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        IReadOnlyList<WordDocumentVisualSnapshot> pages = document.CreateVisualSnapshots();
        Assert.Equal(2, pages.Count);
        for (int index = 0; index < 2; index++) {
            bool marker = even ? index == 0 : index == 1;
            Assert.Equal(marker, pages[index].Drawing.Elements.OfType<OfficeDrawingText>().Any(text => text.Text.Contains("DEFAULTMARK")));
        }
    }
}
