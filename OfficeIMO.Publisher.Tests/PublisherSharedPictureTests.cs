using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public class PublisherSharedPictureTests {
    [Fact]
    public void Shared_asset_placements_keep_independent_crop_geometry_and_selected_replacement_ownership() {
        byte[] png = PublisherPictureEditingFixture.Image(OfficeColor.Red);
        byte[] replacement = PublisherPictureEditingFixture.Image(OfficeColor.White);
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPicture first = document.AddImage(document.Pages[0].Id, png, OfficeLength.Points(72), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
        PublisherImage asset = Assert.Single(document.Images);
        var originalPages = document.Pages;
        PublisherPage master = document.AddMasterPage("Header");
        PublisherPicture second = document.AddImage(master.Id, asset.Id, OfficeLength.Points(24), OfficeLength.Points(24), OfficeLength.Points(48), OfficeLength.Points(48));
        PublisherPage last = document.AddPage(master.Id);
        PublisherPicture third = document.AddImage(last.Id, asset.Id, OfficeLength.Points(216), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
        Assert.Same(asset, Assert.Single(document.Images));
        Assert.Equal(new[] { first.ImageId, first.ImageId }, new[] { second.ImageId, third.ImageId });
        Assert.Single(originalPages[0].Pictures);
        document.SetPictureCrop(second.Id, new OfficeImageSourceCrop(.25, .125, 0, 0));
        document.SetPictureBounds(third.Id, OfficeLength.Points(240), OfficeLength.Points(180), OfficeLength.Points(120), OfficeLength.Points(84));
        document.ReplacePictureImage(first.Id, replacement);
        Assert.Same(asset, document.Images.Single(image => image.Id == asset.Id));
        Assert.Equal(png, Painted(document.MasterPages[0], second.Id).Bytes);
        Assert.Equal(png, Painted(document.Pages[1], third.Id).Bytes);
        Assert.Equal(replacement, Painted(document.Pages[0], first.Id).Bytes);
        Assert.Equal(.25, Assert.Single(document.MasterPages[0].Pictures).Crop.Left);
        Assert.False(first.Crop.HasCrop); Assert.Equal(216, third.X);
        Assert.Equal(240, Assert.Single(document.Pages[1].Pictures).X);
        document.ReplacePictureImage(second.Id, replacement);
        Assert.Contains(document.Images, image => image.Id == asset.Id);
        document.ReplacePictureImage(third.Id, replacement);
        Assert.DoesNotContain(document.Images, image => image.Id == asset.Id);
        Assert.Equal(png, asset.GetBytes());
        Assert.Equal(3, document.Images.Count);
        Assert.Equal(3, document.ReadReport.CurrentObjectCount);
    }

    [Fact]
    public void Reusing_native_asset_keeps_source_bytes_and_failed_admission_does_not_reserve_identifiers() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.pub"));
        PublisherImage asset = Assert.Single(document.Images);
        var pages = document.Pages; var images = document.Images; var report = document.ReadReport;
        uint pageId = pages[0].Id;
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddImage(pageId, 0, OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(40), OfficeLength.Points(40)));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddImage(999, asset.Id, OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(40), OfficeLength.Points(40)));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLength.Points(double.NaN));
        Assert.Throws<OperationCanceledException>(() => document.AddImage(pageId, asset.Id, OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(40), OfficeLength.Points(40), new CancellationToken(true)));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images); Assert.Same(report, document.ReadReport);
        PublisherPicture shared = document.AddImage(pageId, asset.Id, OfficeLength.Points(180), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
        Assert.Equal(asset.Id, shared.ImageId); Assert.Same(asset, Assert.Single(document.Images));
        PublisherPicture independent = document.AddImage(pageId, asset.GetBytes(), OfficeLength.Points(300), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
        Assert.Equal(shared.Id + 1, independent.Id);
        Assert.Equal(asset.Id + 1, independent.ImageId);
        Assert.Equal(2, document.Images.Count);
        Assert.Equal(asset.GetBytes(), document.Images.Single(image => image.Id == independent.ImageId).GetBytes());
        document.ReplacePictureImage(shared.Id, PublisherPictureEditingFixture.Image(OfficeColor.White));
        Assert.Same(asset, document.Images.Single(image => image.Id == asset.Id));
        Assert.Equal(asset.Id, pages[0].Pictures.Single().ImageId);
    }

    private static OfficeDrawingImage Painted(PublisherPage page, uint id) => Assert.Single(PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingImage>(), image => image.SourceElementIds?.Contains("publisher-object-" + id) == true);
}
