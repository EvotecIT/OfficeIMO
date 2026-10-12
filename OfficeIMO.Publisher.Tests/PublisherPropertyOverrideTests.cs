using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPropertyOverrideTests {
    [Theory]
    [InlineData(false, 0U, false, true)]
    [InlineData(true, 0U, false, true)]
    [InlineData(true, 1U, true, true)]
    [InlineData(false, 1U, false, false)]
    public void Picture_reference_clears_and_complex_payloads_do_not_select_an_image_index(bool complex, uint value, bool unsupported, bool blip) {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1), Array.Empty<PublisherBlock>()));
        source.ObjectPages.Add(2, 1);
        var escher = new PublisherEscherData();
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.FromRgb(0, 255, 0)), OfficeImageExportFormat.Png);
        escher.Images.Add(1, new PublisherImage(1, "image/png", image));
        escher.Shapes.Add(new PublisherEscherShape(2, 1, 0, new[] {
            new OfficeArtProperty(0, 0x4104, 1),
            new OfficeArtProperty(1, complex ? (ushort)0x8104 : blip ? (ushort)0x4104 : (ushort)0x0104, value,
                complex ? (int)value : null, complexData: complex ? new byte[value] : null)
        }) { Bounds = new PublisherNativeRectangle(-100 * 12700, -100 * 12700, 100 * 12700, 100 * 12700) });
        PublisherDocument publication = new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
        Assert.DoesNotContain(PublisherNativeTests.Elements(publication.Pages[0].Drawing), item => item is OfficeDrawingImage);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_IMAGE_REFERENCE_UNRESOLVED");
        Assert.Equal(unsupported, publication.ReadReport.FidelityDiagnostics.Any(item => item.Code == "PUB_IMAGE_DATA_UNSUPPORTED"
            && item.Location == "Contents/object/2"));
        Assert.Equal(image, Assert.Single(publication.Images).GetBytes());
    }
}
