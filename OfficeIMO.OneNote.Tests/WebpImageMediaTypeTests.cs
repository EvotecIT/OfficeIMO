namespace OfficeIMO.OneNote.Tests;

public sealed class WebpImageMediaTypeTests {
    [Theory]
    [InlineData("image.webp")]
    [InlineData("image.WEBP")]
    public void ReopenedWebpImageRetainsMediaTypeAndPayload(string fileName) {
        // The native store preserves opaque payloads; format decoding belongs to the image owner.
        byte[] payload = { 1, 2, 3, 4 };
        var section = new OneNoteSection { Name = "Images" };
        var page = new OneNotePage { Title = "WebP" };
        page.DirectContent.Add(new OneNoteImage {
            FileName = fileName,
            MediaType = "image/webp",
            Payload = OneNoteBinaryPayload.FromBytes(payload)
        });
        section.Pages.Add(page);

        using var stream = new MemoryStream(OneNoteSectionWriter.Write(section));
        OneNoteSection loaded = OneNoteSectionReader.Read(stream);
        OneNoteImage image = Assert.IsType<OneNoteImage>(
            Assert.Single(Assert.Single(Assert.Single(loaded.Pages).Outlines).Children));
        Assert.Equal("image/webp", image.MediaType);
        Assert.Equal(fileName, image.FileName);
        Assert.Equal(payload, image.Payload!.ToArray(payload.Length));
    }
}
