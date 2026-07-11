using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Color = OfficeIMO.Drawing.OfficeColor;
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
        using var stream = new MemoryStream(new byte[] {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        });
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

        _ = document.ToDocx();
        await DoTest(document._wordprocessingDocument);
    }
}
