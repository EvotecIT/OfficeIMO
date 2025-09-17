using System.IO;
using System.Threading.Tasks;
using System.Linq;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

/// <summary>
/// Tests removal of images and related cleanup.
/// </summary>
public class ImageRemoveTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task RemoveImageDeletesPart() {
        using var image = new SixLabors.ImageSharp.Image<SixLabors.ImageSharp.PixelFormats.Rgba32>(1, 1);
        using var stream = new MemoryStream();
        image.SaveAsPng(stream);
        stream.Position = 0;
        using var document = WordDocument.Create();
        document.AddHeadersAndFooters();
        var paragraph = document.Header!.Default.AddParagraph();
        paragraph.AddImage(stream, "tiny.png", 50, 50);

        Assert.Single(document.Header!.Default.Images);
        var headerPart = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.First();
        Assert.Single(headerPart.ImageParts);

        document.Header!.Default.Images[0].Remove();

        Assert.Empty(document.Header!.Default.Images);
        Assert.Empty(headerPart.ImageParts);

        document.Save();
        await DoTest(document._wordprocessingDocument);
    }
}
