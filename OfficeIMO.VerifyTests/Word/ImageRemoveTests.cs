using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using SixLabors.ImageSharp;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;
using Assert = OfficeIMO.VerifyTests.TestAssert;

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
        var headers = Assert.NotNull(document.Header);
        var defaultHeader = Assert.IsType<WordHeader>(headers.Default);
        var paragraph = defaultHeader.AddParagraph();
        paragraph.AddImage(stream, "tiny.png", 50, 50);

        Assert.Single(defaultHeader.Images);
        var headerPart = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.First();
        Assert.Single(headerPart.ImageParts);

        defaultHeader.Images[0].Remove();

        Assert.Empty(defaultHeader.Images);
        Assert.Empty(headerPart.ImageParts);

        document.Save();
        await DoTest(document._wordprocessingDocument);
    }
}
