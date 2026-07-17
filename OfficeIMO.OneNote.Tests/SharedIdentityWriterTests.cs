namespace OfficeIMO.OneNote.Tests;

public sealed class SharedIdentityWriterTests {
    [Fact]
    public void IndependentlyEditedTagsSharingLoadedDefinitionReceiveDistinctObjects() {
        var definitionId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var section = new OneNoteSection { Name = "Shared tags" };
        var page = new OneNotePage { Title = "Shared tags" };
        page.DirectContent.Add(TaggedParagraph("First", definitionId));
        page.DirectContent.Add(TaggedParagraph("Second", definitionId));
        section.Pages.Add(page);
        var options = new OneNoteWriterOptions { PreserveUnknownData = false };

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section, options)));
        OneNoteParagraph[] loadedParagraphs = BodyElements<OneNoteParagraph>(loaded);
        Assert.Equal(loadedParagraphs[0].Tags[0].DefinitionId, loadedParagraphs[1].Tags[0].DefinitionId);
        loadedParagraphs[1].Tags[0].Label = "Edited tag";

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(loaded, options)));
        OneNoteParagraph[] result = BodyElements<OneNoteParagraph>(roundTrip);

        Assert.Equal("Shared tag", Assert.Single(result[0].Tags).Label);
        Assert.Equal("Edited tag", Assert.Single(result[1].Tags).Label);
        Assert.NotEqual(result[0].Tags[0].DefinitionId, result[1].Tags[0].DefinitionId);
    }

    [Fact]
    public void IndependentlyEditedAssetsSharingLoadedPayloadReceiveDistinctObjectsAndFileData() {
        var payloadObjectId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        Guid fileDataId = Guid.NewGuid();
        var section = new OneNoteSection { Name = "Shared assets" };
        var page = new OneNotePage { Title = "Shared assets" };
        page.DirectContent.Add(Image("first.png", payloadObjectId, fileDataId, new byte[] { 1, 2, 3 }));
        page.DirectContent.Add(Image("second.png", payloadObjectId, fileDataId, new byte[] { 1, 2, 3 }));
        section.Pages.Add(page);
        var options = new OneNoteWriterOptions { PreserveUnknownData = false };

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section, options)));
        OneNoteImage[] loadedImages = BodyElements<OneNoteImage>(loaded);
        Assert.Equal(loadedImages[0].PayloadObjectId, loadedImages[1].PayloadObjectId);
        Assert.Equal(loadedImages[0].PayloadFileDataId, loadedImages[1].PayloadFileDataId);
        loadedImages[1].Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 4, 5, 6 });

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(loaded, options)));
        OneNoteImage[] result = BodyElements<OneNoteImage>(roundTrip);

        Assert.Equal(new byte[] { 1, 2, 3 }, result[0].Payload!.ToArray(16));
        Assert.Equal(new byte[] { 4, 5, 6 }, result[1].Payload!.ToArray(16));
        Assert.NotEqual(result[0].PayloadObjectId, result[1].PayloadObjectId);
        Assert.NotEqual(result[0].PayloadFileDataId, result[1].PayloadFileDataId);
    }

    private static OneNoteParagraph TaggedParagraph(string text, OneNoteExtendedGuid definitionId) {
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = text });
        paragraph.Tags.Add(new OneNoteTag {
            DefinitionId = definitionId,
            ActionItemType = 0,
            Label = "Shared tag",
            Shape = 13,
            IsCheckable = false
        });
        return paragraph;
    }

    private static OneNoteImage Image(
        string fileName,
        OneNoteExtendedGuid payloadObjectId,
        Guid fileDataId,
        byte[] payload) => new OneNoteImage {
            FileName = fileName,
            Payload = OneNoteBinaryPayload.FromBytes(payload),
            PayloadObjectId = payloadObjectId,
            PayloadFileDataId = fileDataId,
            PayloadFileExtension = ".png"
        };

    private static T[] BodyElements<T>(OneNoteSection section) where T : OneNoteElement {
        OneNotePage page = Assert.Single(section.Pages);
        Assert.Empty(page.DirectContent);
        return Assert.Single(page.Outlines).Children.OfType<T>().ToArray();
    }
}
