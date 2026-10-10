using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCustomPicturePathTests {
    [Theory]
    [InlineData(4)]
    [InlineData(0xFFF0)]
    public void Compact_picture_masks_keep_high_coordinate_vertices_in_the_visible_frame(int elementSize) {
        byte[] original = PublisherPictureEffectTests.PictureInput(new());
        byte[] input = PublisherDrawingFixture.Mutate(original, new() {
            [0x140] = 0, [0x141] = 0, [0x142] = 60000, [0x143] = 60000, [0x144] = 1
        }, 0, new() { [0x145] = PublisherCustomPathTests.Vertices(new[] { (20000, 20000), (50000, 20000), (20000, 50000) }, elementSize) }, 362);
        PublisherDocument publication = PublisherDocument.Load(input, new PublisherReadOptions { ImageCodec = new SolidCodec() });
        OfficeDrawingGroup group = publication.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing))
            .OfType<OfficeDrawingGroup>().Single(item => item.SourceElementIds?.Contains("publisher-object-362") == true);
        Assert.Equal(group.ClipPath.Width * 5 / 6, group.ClipPath.Commands[1].Point.X, 6);
        Assert.Equal(group.ClipPath.Height * 5 / 6, group.ClipPath.Commands[2].Point.Y, 6);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(group.ClipPath.Width, group.ClipPath.Height)
            .AddClippedDrawing(group.InnerDrawing, 0, 0, group.ClipPath));
        Assert.Equal(OfficeColor.Red, raster.GetPixel((int)(raster.Width * 0.4), (int)(raster.Height * 0.4)));
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.8), (int)(raster.Height * 0.8)).A);
    }

    [Fact]
    public void Custom_picture_mask_keeps_inset_coordinates_and_original_assets() {
        byte[] original = PublisherPictureEffectTests.PictureInput(new());
        byte[] input = PublisherDrawingFixture.Mutate(original, new() {
            [0x140] = 0, [0x141] = 0, [0x142] = 100, [0x143] = 100, [0x144] = 1
        }, 0, new() { [0x145] = PublisherCustomPathTests.Vertices(new[] { (25, 25), (75, 25), (25, 75) }) }, 362);
        var options = new PublisherReadOptions { ImageCodec = new SolidCodec() };
        PublisherDocument publication = PublisherDocument.Load(input, options);
        PublisherDocument control = PublisherDocument.Load(original, options);
        Assert.Equal(control.Images.SelectMany(image => image.GetBytes()), publication.Images.SelectMany(image => image.GetBytes()));
        OfficeDrawingGroup group = publication.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing))
            .OfType<OfficeDrawingGroup>().Single(item => item.SourceElementIds?.Contains("publisher-object-362") == true);
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Equal(new OfficePoint(group.ClipPath.Width / 4, group.ClipPath.Height / 4), group.ClipPath.Commands[0].Point);
        Assert.Contains("publisher-object-362", Assert.Single(group.InnerDrawing.Images).SourceElementIds!);
        var drawing = new OfficeDrawing(group.ClipPath.Width, group.ClipPath.Height)
            .AddClippedDrawing(group.InnerDrawing, 0, 0, group.ClipPath);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.1), (int)(raster.Height * 0.1)).A);
        Assert.Equal(OfficeColor.Red, raster.GetPixel((int)(raster.Width * 0.3), (int)(raster.Height * 0.3)));
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.8), (int)(raster.Height * 0.8)).A);
    }

    private sealed class SolidCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = new OfficeRasterImage(20, 20, OfficeColor.Red); return true;
        }
    }
}
