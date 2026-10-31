using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPictureCropTests {
    private static readonly OfficeImageSourceCrop Edges = OfficeImageSourceCrop.FromStrictFractions(0.25, 0.125, 0.0625, 0.1875);

    [Theory]
    [InlineData("set")]
    [InlineData("clear")]
    [InlineData("replace-and-resize")]
    public void Native_crop_saves_retain_other_placements_assets_and_the_selected_frame_contract(string operation) {
        string fixture = operation == "clear" ? "PictureCropped" : "PictureOwnership";
        var document = PublisherDocument.Load(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/native-{fixture.ToLowerInvariant()}.pub"));
        PublisherPicture original = document.Pages[0].Pictures.Single(picture => picture.Id == 298);
        PublisherPicture duplicate = document.Pages[0].Pictures.Single(picture => picture.Id == 300);
        PublisherPicture master = document.MasterPages.Single().Pictures.Single();
        byte[] image = document.Images.Single(asset => asset.Id == original.ImageId).GetBytes();
        OfficeImageSourceCrop crop = operation == "clear" ? default : Edges;
        document.SetPictureCrop(original.Id, crop);
        byte[]? replacement = null;
        if (operation == "replace-and-resize") {
            replacement = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/PictureReplacement/native-picturereplacement-replacement.png"));
            document.ReplacePictureImage(original.Id, replacement);
            document.SetPictureBounds(original.Id, 84, 132, 120, 84);
        }
        var current = document.Pages[0].Pictures.Single(picture => picture.Id == original.Id);
        Assert.Equal(crop, current.Crop);
        PublisherDocument reopened = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherPicture actual = reopened.Pages[0].Pictures.Single(picture => picture.Id == original.Id);
        Assert.Equal(crop, actual.Crop);
        Assert.Equal((current.Id, current.PageId, current.X, current.Y, current.Width, current.Height),
            (actual.Id, actual.PageId, actual.X, actual.Y, actual.Width, actual.Height));
        Assert.Equal(replacement ?? image, reopened.Images.Single(asset => asset.Id == actual.ImageId).GetBytes());
        Assert.Equal(image, reopened.Images.Single(asset => asset.Id == original.ImageId).GetBytes());
        foreach (PublisherPicture before in new[] { duplicate, master }) {
            PublisherPicture after = reopened.Pages.Concat(reopened.MasterPages).SelectMany(page => page.Pictures).Single(picture => picture.Id == before.Id);
            Assert.Equal((before.ImageId, before.PageId, before.X, before.Y, before.Width, before.Height, before.Crop),
                (after.ImageId, after.PageId, after.X, after.Y, after.Width, after.Height, after.Crop));
        }
    }

    [Fact]
    public void Source_free_crop_updates_the_scene_without_mutating_assets_neighbors_or_old_snapshots() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.png"));
        PublisherPicture original = document.AddPicture(document.Pages[0].Id, bytes, 72, 144, 96, 96);
        PublisherPicture neighbor = document.AddPicture(document.Pages[0].Id, bytes, 216, 144, 96, 96);
        PublisherPage oldPage = document.Pages[0];
        string oldScene = OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing());
        document.SetPictureCrop(original.Id, Edges);
        PublisherPage current = document.Pages[0];
        Assert.False(original.Crop.HasCrop);
        Assert.Equal(Edges, current.Pictures.Single(picture => picture.Id == original.Id).Crop);
        Assert.Equal(neighbor.Crop, current.Pictures.Single(picture => picture.Id == neighbor.Id).Crop);
        OfficeDrawingImage image = Assert.Single(PublisherNativeTests.Elements(current.CreateDrawing()).OfType<OfficeDrawingImage>(),
            item => item.SourceElementIds?.Contains("publisher-object-" + original.Id) == true);
        Assert.Equal(Edges, image.Projection.SourceCrop);
        Assert.All(document.Images, asset => Assert.Equal(bytes, asset.GetBytes()));
        Assert.Equal(oldScene, OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing()));
        Assert.NotEqual(oldScene, document.ToSvg());
    }

    [Fact]
    public void Invalid_crop_unknown_identity_and_cancellation_preserve_the_current_state() {
        PublisherDocument document = PublisherPictureEditingFixture.Create();
        var pages = document.Pages; var images = document.Images; string svg = document.ToSvg();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetPictureCrop(101, new OfficeImageSourceCrop(0.75, 0, 0.25, 0)));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetPictureCrop(uint.MaxValue, Edges));
        Assert.Throws<OperationCanceledException>(() => document.SetPictureCrop(101, Edges, new CancellationToken(true)));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images); Assert.Equal(svg, document.ToSvg());
    }

    [Fact]
    public void Native_crop_quantization_is_reported_without_changing_the_requested_model() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.pub"));
        uint id = document.Pages[0].Pictures.Single().Id;
        var requested = new OfficeImageSourceCrop(0.1, 0.2, 0.15, 0.05);
        document.SetPictureCrop(id, requested);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_NATIVE_PICTURE_CROP_QUANTIZED"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        PublisherDocument reopened = PublisherDocument.Load(result.Value);
        Assert.Equal(new OfficeImageSourceCrop(6554D / 65536, 13107D / 65536, 9830D / 65536, 3277D / 65536), reopened.Pages[0].Pictures.Single().Crop);
        Assert.Equal(requested, document.Pages[0].Pictures.Single().Crop);
    }

    [Fact]
    public void Native_crop_rounding_that_collapses_the_visible_area_refuses_output() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.pub"));
        document.SetPictureCrop(document.Pages[0].Pictures.Single().Id, new OfficeImageSourceCrop(0.4999999, 0, 0.4999999, 0));
        var exception = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(exception.Report.FidelityDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure
            && diagnostic.Message.Contains("16.16"));
    }
}
