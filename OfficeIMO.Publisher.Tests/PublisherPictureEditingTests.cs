using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPictureEditingTests {
    [Fact]
    public void Native_picture_replacement_and_bounds_reach_current_svg_and_pdf_without_mutating_snapshots() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.pub"));
        PublisherPage oldPage = Assert.Single(document.Pages);
        PublisherPicture original = Assert.Single(oldPage.Pictures);
        Assert.True(original.CanEdit, original.EditLimitation);
        Assert.Equal(72, original.X, 5); Assert.Equal(72, original.Y, 5);
        Assert.Equal(96, original.Width, 5); Assert.Equal(96, original.Height, 5);
        PublisherImage originalAsset = Assert.Single(document.Images);
        byte[] originalBytes = originalAsset.GetBytes();
        var sourceReport = document.SourceReadReport;
        byte[] replacement = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(0, 255, 0));
        byte[] expected = (byte[])replacement.Clone();
        string before = document.ToSvg();
        string oldDrawing = OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing());
        document.ReplacePictureImage(original.Id, replacement);
        Array.Clear(replacement, 0, replacement.Length);
        document.SetPictureBounds(original.Id, 36, 48, 120, 84);
        PublisherPicture current = Assert.Single(document.Pages[0].Pictures);
        Assert.Equal(original.Id, current.Id); Assert.Equal(original.PageId, current.PageId);
        Assert.Equal(original.Crop, current.Crop);
        Assert.Equal(36, current.X); Assert.Equal(48, current.Y);
        Assert.Equal(120, current.Width); Assert.Equal(84, current.Height);
        Assert.Equal(new OfficePoint(36, 48), current.PageTransform.TransformPoint(new OfficePoint(0, 0)));
        Assert.NotEqual(original.ImageId, current.ImageId);
        Assert.Equal(expected, document.Images.Single(image => image.Id == current.ImageId).GetBytes());
        Assert.Equal(originalBytes, originalAsset.GetBytes());
        Assert.Equal(72, original.X, 5); Assert.Equal(originalAsset.Id, original.ImageId);
        Assert.Equal(oldDrawing, OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing()));
        Assert.Contains(Convert.ToBase64String(expected), document.ToSvg());
        Assert.NotEqual(before, document.ToSvg());
        Assert.Same(sourceReport, document.SourceReadReport);
        byte[] pdf = document.ToPdfBytes();
        PdfReadDocument readPdf = PdfReadDocument.Open(pdf);
        PdfImagePlacement placement = Assert.Single(Assert.Single(readPdf.Pages).GetImagePlacements());
        Assert.Equal(36, placement.X, 4); Assert.Equal(120, placement.Width, 4); Assert.Equal(84, placement.Height, 4);
        Assert.Single(readPdf.ExtractImages());
        OfficeDrawing copy = document.Pages[0].CreateDrawing();
        copy.AddShape(OfficeShape.Rectangle(10, 10), 0, 0);
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingShape>(), shape => shape.Shape.Width == 10);
    }

    [Fact]
    public void Replacement_changes_one_shared_placement_and_master_edits_reach_each_referring_page() {
        PublisherDocument document = PublisherPictureEditingFixture.Create();
        PublisherPage oldPage = document.Pages[0], oldMaster = Assert.Single(document.MasterPages);
        PublisherPicture[] oldPictures = oldPage.Pictures.ToArray();
        string[] paintOrder = PublisherNativeTests.Elements(oldPage.CreateDrawing()).OfType<OfficeDrawingImage>()
            .Select(image => image.SourceElementIds!.Single()).ToArray();
        Assert.Equal(new uint[] { 101, 102 }, oldPictures.Select(picture => picture.Id));
        Assert.All(oldPictures, picture => Assert.Equal(1, picture.ImageId));
        byte[] green = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(0, 255, 0));
        document.ReplacePictureImage(101, green);
        Assert.NotEqual(1, document.Pages[0].Pictures[0].ImageId);
        Assert.Equal(1, document.Pages[0].Pictures[1].ImageId);
        Assert.Equal(1, document.MasterPages[0].Pictures[0].ImageId);
        Assert.Equal(green, Painted(document.Pages[0], 101).Bytes);
        Assert.NotEqual(green, Painted(document.Pages[0], 102).Bytes);
        Assert.Equal(new uint[] { 101, 102 }, document.Pages[0].Pictures.Select(picture => picture.Id));
        Assert.Equal(paintOrder, PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingImage>()
            .Select(image => image.SourceElementIds!.Single()));
        byte[] blue = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(0, 0, 255));
        document.ReplacePictureImage(103, blue);
        document.SetPictureBounds(103, 24, 160, 90, 60);
        PublisherPicture masterPicture = Assert.Single(document.MasterPages[0].Pictures);
        Assert.Equal(20U, masterPicture.PageId);
        Assert.All(document.Pages, page => {
            Assert.Equal(20U, page.MasterPageId);
            Assert.DoesNotContain(page.Pictures, picture => picture.Id == 103);
            Assert.Equal(blue, Painted(page, 103).Bytes);
            Assert.Equal(24, Painted(page, 103).Projection.Placement.X);
            Assert.Equal(160, Painted(page, 103).Projection.Placement.Y);
        });
        Assert.Equal(1, oldMaster.Pictures[0].ImageId);
        Assert.All(oldPictures, picture => Assert.Equal(1, picture.ImageId));
        Assert.Equal(300, document.Pages[0].Width); Assert.Equal(300, document.MasterPages[0].Height);
        Assert.Contains(Convert.ToBase64String(blue), document.ToSvg(0));
        Assert.Contains(Convert.ToBase64String(blue), document.ToSvg(1));
        Assert.Equal(2, PdfReadDocument.Open(document.ToPdfBytes()).Pages.Count);
    }

    [Theory]
    [InlineData("effects")]
    [InlineData("hidden")]
    [InlineData("rotation")]
    [InlineData("reflected")]
    [InlineData("grouped")]
    [InlineData("nonrectangular")]
    [InlineData("unresolved")]
    [InlineData("no-anchor")]
    [InlineData("invalid-crop")]
    public void Unsupported_picture_edits_are_explicit_and_leave_readable_state_unchanged(string restriction) {
        PublisherDocument document = PublisherPictureEditingFixture.Create(restriction: restriction);
        PublisherPicture picture = document.Pages[0].Pictures.Single(item => item.Id == 101);
        Assert.False(picture.CanEdit); Assert.False(string.IsNullOrEmpty(picture.EditLimitation));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_PICTURE_EDIT_UNSUPPORTED"
            && diagnostic.Location == "Contents/object/101" && diagnostic.Message == picture.EditLimitation);
        var pages = document.Pages; var images = document.Images; string before = document.ToSvg();
        Assert.Throws<NotSupportedException>(() => document.SetPictureBounds(101, 10, 10, 50, 50));
        Assert.Throws<NotSupportedException>(() => document.ReplacePictureImage(101, PublisherPictureEditingFixture.Image(OfficeColor.Black)));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images); Assert.Equal(before, document.ToSvg());
        Assert.NotNull(Painted(document.Pages[0], 102));
    }

    [Fact]
    public void Invalid_content_geometry_unknown_identity_and_cancellation_leave_current_picture_state_usable() {
        PublisherDocument document = PublisherPictureEditingFixture.Create();
        var pages = document.Pages; var images = document.Images; var report = document.ReadReport;
        string before = document.ToSvg();
        Assert.Throws<ArgumentNullException>(() => document.ReplacePictureImage(101, null!));
        Assert.Throws<ArgumentException>(() => document.ReplacePictureImage(101, new byte[] { 1, 2, 3 }));
        byte[] truncated = PublisherPictureEditingFixture.Image(OfficeColor.Black).Take(33).ToArray();
        Assert.Throws<ArgumentException>(() => document.ReplacePictureImage(101, truncated));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetPictureBounds(101, 1, 1, 0, 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetPictureBounds(101, double.NaN, 1, 20, 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetPictureBounds(999, 1, 1, 20, 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ReplacePictureImage(999, PublisherPictureEditingFixture.Image(OfficeColor.Black)));
        using var source = new CancellationTokenSource(); source.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.SetPictureBounds(101, 1, 1, 20, 20, source.Token));
        Assert.Throws<OperationCanceledException>(() => document.ReplacePictureImage(101, PublisherPictureEditingFixture.Image(OfficeColor.Black), source.Token));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images); Assert.Same(report, document.ReadReport);
        Assert.Equal(before, document.ToSvg());
        document.SetPictureBounds(101, -12, 40, 60, 60);
        Assert.Equal(-12, document.Pages[0].Pictures[0].X);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_PAGE_BLEED_CLIPPED"
            && diagnostic.Location == "Contents/object/101");
    }

    [Fact]
    public void Replacement_reserves_missing_native_asset_slots_and_does_not_resolve_an_unrelated_picture() {
        PublisherDocument document = PublisherPictureEditingFixture.Create(restriction: "unresolved-sibling");
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingImage>(),
            image => image.SourceElementIds?.Contains("publisher-object-102") == true);
        byte[] replacement = PublisherPictureEditingFixture.Image(OfficeColor.Black);
        document.ReplacePictureImage(101, replacement);
        int imageId = document.Pages[0].Pictures.Single(picture => picture.Id == 101).ImageId!.Value;
        Assert.DoesNotContain(imageId, new[] { 1, 2, 3, 4 });
        Assert.Equal(replacement, Painted(document.Pages[0], 101).Bytes);
        Assert.Equal(2, document.Pages[0].Pictures.Single(picture => picture.Id == 102).ImageId);
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingImage>(),
            image => image.SourceElementIds?.Contains("publisher-object-102") == true);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_IMAGE_REFERENCE_UNRESOLVED"
            && diagnostic.Location == "Contents/object/102");
    }

    [Fact]
    public void Picture_movement_recalculates_existing_story_wrap_from_current_semantic_bounds() {
        PublisherDocument document = PublisherPictureEditingFixture.Create(withWrapping: true);
        string story = Assert.Single(document.TextStories).Text;
        OfficeDrawingRichText[] Regions() => PublisherNativeTests.Elements(document.Pages[0].CreateDrawing())
            .OfType<OfficeDrawingRichText>().Where(region => region.SourceElementIds?.Contains("publisher-object-201") == true).ToArray();
        OfficeDrawingRichText[] before = Regions();
        Assert.True(before.Length > 1);
        document.SetPictureBounds(101, 280, 30, 72, 72);
        OfficeDrawingRichText after = Assert.Single(Regions());
        Assert.True(after.Width > before.Min(region => region.Width));
        Assert.Equal(story, Assert.Single(document.TextStories).Text);
        Assert.Equal(new uint[] { 101 }, Assert.Single(document.Pages[0].TextFrames).WrapObjectIds);
        Assert.Equal(280, document.Pages[0].Pictures[0].X);
    }

    [Fact]
    public void Replacement_obeys_pixel_and_projection_budgets_and_releases_superseded_edit_assets() {
        PublisherDocument document = PublisherPictureEditingFixture.Create(new PublisherReadOptions { MaximumRasterPixels = 4 });
        var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.ReplacePictureImage(101, PublisherPictureEditingFixture.Image(OfficeColor.Black, 3, 2)));
        Assert.Same(pages, document.Pages);
        document.ReplacePictureImage(101, PublisherPictureEditingFixture.Image(OfficeColor.Black));
        int oldEditId = document.Pages[0].Pictures[0].ImageId!.Value;
        document.ReplacePictureImage(101, PublisherPictureEditingFixture.Image(OfficeColor.White));
        Assert.DoesNotContain(document.Images, image => image.Id == oldEditId);
        Assert.Contains(document.Images, image => image.Id == 1);
        Assert.Equal(2, document.Images.Count);
        Assert.Equal(1, document.Pages[0].Pictures[1].ImageId);

        int sourceBytes = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(255, 0, 0)).Length;
        PublisherDocument bounded = PublisherPictureEditingFixture.Create(new PublisherReadOptions { MaximumTotalImageBytes = sourceBytes * 12 });
        byte[] padded = PublisherPictureWorkFixture.PaddedPng();
        var boundedPages = bounded.Pages; string before = bounded.ToSvg();
        Assert.ThrowsAny<ArgumentException>(() => bounded.ReplacePictureImage(101, padded));
        Assert.Same(boundedPages, bounded.Pages); Assert.Equal(before, bounded.ToSvg());

        PublisherDocument projectionBounded = PublisherPictureEditingFixture.Create(new PublisherReadOptions { MaximumTotalImageBytes = sourceBytes * 9 });
        var projectionPages = projectionBounded.Pages; var projectionImages = projectionBounded.Images;
        string projectionBefore = projectionBounded.ToSvg();
        Assert.Throws<InvalidDataException>(() => projectionBounded.ReplacePictureImage(101, PublisherPictureEditingFixture.Image(OfficeColor.White)));
        Assert.Same(projectionPages, projectionBounded.Pages); Assert.Same(projectionImages, projectionBounded.Images);
        Assert.Equal(projectionBefore, projectionBounded.ToSvg());
    }

    private static OfficeDrawingImage Painted(PublisherPage page, uint id) => PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingImage>().Single(image => image.SourceElementIds?.Contains("publisher-object-" + id) == true);
}
