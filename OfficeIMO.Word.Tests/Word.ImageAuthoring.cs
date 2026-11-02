using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordImageAuthoringTests {
    private static byte[] Gif => Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");

    [Fact]
    public void ImageHyperlinkKeepsNeighboringTextOutsideTheLink() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Before");
        using var source = new MemoryStream(Gif);
        var image = paragraph.AddImage(source, "sample.gif", OfficeIMO.OfficeLength.Pixels(32), OfficeIMO.OfficeLength.Pixels(32));
        paragraph.AddText("After");
        image.Hyperlink = new Uri("https://example.com/image");
        image.SetMaximumWidth(OfficeLength.Pixels(16));

        using var stream = new MemoryStream(document.ToBytes());
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        var link = Assert.Single(main.Document!.Descendants<Hyperlink>());
        Assert.Empty(link.Descendants<Text>());
        Assert.Single(link.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.Equal(new[] { "Before", "After" }, main.Document!.Descendants<Text>().Select(text => text.Text));
        Assert.Equal(new Uri("https://example.com/image"), Assert.Single(main.HyperlinkRelationships).Uri);
        Assert.Equal(16d, (image.Width?.ToPixels()));
        Assert.Equal(32d, (image.Height?.ToPixels()));
        Assert.True(source.CanRead);
    }

    [Fact]
    public void DocumentImageOptionsPreserveMetadataAndEmbeddedReading() {
        using var document = WordDocument.Create();
        using var source = new MemoryStream(Gif);
        var image = document.AddImage(source, OfficeLength.Pixels(32), OfficeLength.Pixels(16), new WordImageOptions {
            FileName = "sample.gif", Title = "Sample", Description = "An embedded sample",
            MaximumWidth = OfficeLength.Pixels(24), Alignment = WordParagraphAlignment.Center
        });
        Assert.Equal("Sample", image.Title);
        Assert.Equal("An embedded sample", image.Description);
        Assert.Equal(WordParagraphAlignment.Center, image.Paragraph.ParagraphAlignment);
        Assert.Equal(24d, (image.Width?.ToPixels()));
        Assert.Equal(Gif, Assert.Single(document.GetImageBytes()));
        foreach (var embedded in document.OpenImageStreams()) {
            using (embedded) Assert.Equal(Gif.Length, embedded.Length);
        }
        Assert.True(source.CanRead);
    }
}
