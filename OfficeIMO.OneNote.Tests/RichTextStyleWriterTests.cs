namespace OfficeIMO.OneNote.Tests;

public sealed class RichTextStyleWriterTests {
    [Fact]
    public void IndependentlyEditedRunsSharingLoadedStyleReceiveDistinctObjects() {
        var sharedStyleId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var first = new OneNoteTextRun { Text = "first" };
        first.Style.Bold = true;
        first.StyleObjectId = sharedStyleId;
        var second = new OneNoteTextRun { Text = "second" };
        second.Style.Bold = true;
        second.StyleObjectId = sharedStyleId;

        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(first);
        paragraph.Runs.Add(second);
        var page = new OneNotePage { Title = "Shared styles" };
        page.DirectContent.Add(paragraph);
        var section = new OneNoteSection { Name = "Styles" };
        section.Pages.Add(page);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteParagraph loadedParagraph = BodyParagraph(Assert.Single(loaded.Pages));
        Assert.Equal(loadedParagraph.Runs[0].StyleObjectId, loadedParagraph.Runs[1].StyleObjectId);
        loadedParagraph.Runs[0].Style.Italic = true;

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(loaded).ObjectSpaces[1];
        OneNoteWriteObject richText = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidRichTextNode &&
            item.Properties.Any(property => (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.TextRunFormatting));
        OneNoteExtendedGuid[] styleIds = Assert.Single(richText.Properties,
            property => (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.TextRunFormatting).References.ToArray();
        Assert.Equal(2, styleIds.Length);
        Assert.NotEqual(styleIds[0], styleIds[1]);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(loaded)));
        OneNoteParagraph result = BodyParagraph(Assert.Single(roundTrip.Pages));
        Assert.Collection(result.Runs,
            run => {
                Assert.True(run.Style.Bold);
                Assert.True(run.Style.Italic);
            },
            run => {
                Assert.True(run.Style.Bold);
                Assert.Null(run.Style.Italic);
            });
    }

    private static OneNoteParagraph BodyParagraph(OneNotePage page) {
        Assert.Empty(page.DirectContent);
        return Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(page.Outlines).Children));
    }
}
