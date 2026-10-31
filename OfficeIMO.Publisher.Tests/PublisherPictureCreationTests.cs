using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPictureCreationTests {
    [Fact]
    public void New_placements_keep_native_shared_assets_and_detached_snapshots_while_replacement_prunes_only_authored_bytes() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/PictureOwnership/native-pictureownership.pub"));
        PublisherPage oldPage = document.Pages[1]; PublisherPage oldMaster = Assert.Single(document.MasterPages);
        PublisherImage nativeAsset = Assert.Single(document.Images); byte[] original = nativeAsset.GetBytes();
        var source = document.SourceReadReport; int sourceObjects = document.ReadReport.SourceObjectCount, records = document.ReadReport.InspectedRecordCount;
        PublisherPicture first = document.AddPicture(oldPage.Id, original, 40, 80, 96, 96);
        PublisherImage oldAuthoredAsset = document.Images.Single(asset => asset.Id == first.ImageId);
        PublisherPicture second = document.AddPicture(oldPage.Id, original, 160, 80, 96, 96);
        Assert.NotEqual(first.ImageId, second.ImageId); Assert.NotEqual(nativeAsset.Id, first.ImageId);
        var pictures = document.Pages[1].Pictures;
        byte[] changed = PublisherPictureEditingFixture.Image(OfficeColor.White);
        document.ReplacePictureImage(first.Id, changed);
        document.SetPictureBounds(first.Id, -8, 90, 120, 84);
        PublisherPicture current = document.Pages[1].Pictures.Single(picture => picture.Id == first.Id);
        Assert.Equal(first.Id, current.Id); Assert.Equal(first.PageId, current.PageId); Assert.Equal(first.Crop, current.Crop);
        Assert.True(current.CanEdit); Assert.Null(current.EditLimitation); Assert.True(current.HasPlacement);
        Assert.False(current.Crop.HasCrop); Assert.Equal(-8, current.X); Assert.Equal(90, current.Y);
        Assert.Equal(120, current.Width); Assert.Equal(84, current.Height);
        Assert.Equal(OfficeTransform.Translate(-8, 90), current.PageTransform);
        Assert.DoesNotContain(document.Images, asset => asset.Id == first.ImageId);
        Assert.Equal(original, oldAuthoredAsset.GetBytes()); Assert.Equal(original, nativeAsset.GetBytes());
        Assert.Equal(original, document.Images.Single(asset => asset.Id == second.ImageId).GetBytes());
        Assert.Equal(changed, Painted(document.Pages[1], current.Id).Bytes);
        Assert.Same(nativeAsset, document.Images.Single(asset => asset.Id == nativeAsset.Id));
        Assert.All(document.Pages[0].Pictures.Concat(document.MasterPages[0].Pictures), picture => Assert.Equal(nativeAsset.Id, picture.ImageId));
        Assert.Equal(2, document.Pages[0].Pictures.Count); Assert.Single(document.MasterPages[0].Pictures);
        Assert.Empty(oldPage.Pictures); Assert.Single(oldMaster.Pictures); Assert.Equal(first.ImageId, pictures[0].ImageId); Assert.Equal(40, first.X);
        Assert.Equal(new[] { first.Id, second.Id }, document.Pages[1].Pictures.Select(picture => picture.Id));
        Assert.Equal(sourceObjects + 2, document.ReadReport.CurrentObjectCount); Assert.Equal(sourceObjects + 2, document.ReadReport.ProjectedObjectCount);
        Assert.Equal(sourceObjects, document.ReadReport.SourceObjectCount); Assert.Equal(records, document.ReadReport.InspectedRecordCount);
        Assert.Same(source, document.SourceReadReport); Assert.True(document.HasNativeSource); Assert.All(document.Pages, page => Assert.True(page.HasNativeOrigin));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_PAGE_BLEED_CLIPPED" && diagnostic.Location == "publication/picture/" + current.Id);
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.Equal(new[] { 3, 3 }, pdf.Pages.Select(page => page.GetImagePlacements().Count));
    }

    [Fact]
    public void Addition_reserves_unavailable_native_image_slots_without_resolving_the_missing_sibling() {
        PublisherDocument document = PublisherPictureEditingFixture.Create(restriction: "unresolved-sibling");
        PublisherPage original = document.Pages[0];
        PublisherPicture added = document.AddPicture(original.Id, Png(), 20, 220, 40, 40);
        Assert.True(added.ImageId > 4);
        Assert.Equal(2, document.Pages[0].Pictures.Single(picture => picture.Id == 102).ImageId);
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingImage>(),
            image => image.SourceElementIds?.Contains("publisher-object-102") == true);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_IMAGE_REFERENCE_UNRESOLVED" && item.Location == "Contents/object/102");
        Assert.Equal(new uint[] { 101, 102 }, original.Pictures.Select(picture => picture.Id));
        Assert.Equal(new uint[] { 101, 102, added.Id }, document.Pages[0].Pictures.Select(picture => picture.Id));
    }

    [Theory]
    [InlineData("animated")]
    [InlineData("oriented")]
    public void Creation_rejects_static_projection_loss_without_narrowing_existing_replacement_support(string kind) {
        byte[] source = kind == "animated" ? AnimatedPng() : OrientedJpeg();
        Assert.True(OfficeRasterImageDecoder.TryDecode(source, new OfficeRasterDecodeOptions(), out _, out OfficeRasterDecodeInfo info));
        Assert.True(kind == "animated" ? info.IsAnimated && info.FrameCount == 2 : info.SourceOrientation == OfficeImageOrientation.Rotate90Clockwise);
        PublisherDocument document = PublisherDocument.Create(300, 400);
        uint pageId = document.Pages[0].Id;
        var pages = document.Pages; var images = document.Images;
        ArgumentException failure = Assert.Throws<ArgumentException>(() => document.AddPicture(pageId, source, 10, 10, 40, 30));
        Assert.Contains("single static PNG or JPEG with normal encoded orientation", failure.Message);
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images);
        PublisherPicture first = document.AddPicture(pageId, Png(), 10, 10, 40, 30);
        Assert.Equal(2U, first.Id); Assert.Equal(1, first.ImageId);
        // Replacement has an established PNG/JPEG contract independent of the
        // narrower new-placement policy. Retain those exact encoded bytes.
        document.ReplacePictureImage(first.Id, source);
        Assert.Equal(source, Assert.Single(document.Images).GetBytes());
        Assert.Equal(first.Id, Assert.Single(document.Pages[0].Pictures).Id);
        Assert.Equal(1, document.ReadReport.CurrentObjectCount);
        pages = document.Pages; images = document.Images;
        failure = Assert.Throws<ArgumentException>(() => document.AddPicture(pageId, Assert.Single(images).Id, 60, 10, 40, 30));
        Assert.Contains("single static PNG or JPEG with normal encoded orientation", failure.Message);
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images);
    }

    [Fact]
    public void Invalid_geometry_content_owner_and_cancellation_leave_both_allocators_and_snapshots_unchanged() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        uint pageId = document.Pages[0].Id; byte[] png = Png();
        var pages = document.Pages; var images = document.Images; var report = document.ReadReport; string before = document.ToSvg();
        Assert.Throws<ArgumentNullException>(() => document.AddPicture(pageId, null!, 10, 10, 40, 40));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPicture(999, png, 10, 10, 40, 40));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPicture(pageId, png, double.NaN, 10, 40, 40));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPicture(pageId, png, 10, 10, 0, 40));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPicture(pageId, png, double.MaxValue, 10, double.MaxValue, 40));
        Assert.Throws<ArgumentException>(() => document.AddPicture(pageId, Array.Empty<byte>(), 10, 10, 40, 40));
        Assert.Throws<ArgumentException>(() => document.AddPicture(pageId, new byte[] { 1, 2, 3 }, 10, 10, 40, 40));
        Assert.Throws<ArgumentException>(() => document.AddPicture(pageId, png.Take(33).ToArray(), 10, 10, 40, 40));
        Assert.Throws<OperationCanceledException>(() => document.AddPicture(pageId, png, 10, 10, 40, 40, new CancellationToken(true)));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
        PublisherPicture added = document.AddPicture(pageId, png, 10, 10, 40, 40);
        Assert.Equal(2U, added.Id); Assert.Equal(1, added.ImageId);
        byte[] detached = Assert.Single(document.Images).GetBytes(); detached[0] = 0;
        Assert.Equal(png, Assert.Single(document.Images).GetBytes());

        PublisherDocument unqualified = PublisherShapeEditingFixture.Create(restriction: "unqualified-master");
        var unchanged = unqualified.Pages;
        Assert.Throws<NotSupportedException>(() => unqualified.AddPicture(unqualified.MasterPages[0].Id, png, 10, 10, 40, 40));
        Assert.Same(unchanged, unqualified.Pages);
        PublisherDocument bounded = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.pub"));
        Assert.Throws<ArgumentOutOfRangeException>(() => bounded.AddPicture(bounded.Pages[0].Id, png, 1_000_000, 0, 40, 40));
    }

    [Theory]
    [InlineData("encoded-bytes")]
    [InlineData("decoded-pixels")]
    [InlineData("retained-bytes")]
    public void Configured_image_admission_bounds_reject_addition_without_publishing_an_asset(string bound) {
        byte[] png = Png();
        var options = new PublisherReadOptions();
        if (bound == "encoded-bytes") options.MaximumImageBytes = png.Length - 1;
        if (bound == "decoded-pixels") options.MaximumRasterPixels = 3;
        if (bound == "retained-bytes") options.MaximumTotalImageBytes = png.Length - 1;
        PublisherDocument document = PublisherDocument.Create(300, 400, options);
        var pages = document.Pages; var report = document.ReadReport;
        ArgumentException failure = Assert.Throws<ArgumentException>(() => document.AddPicture(pages[0].Id, png, 10, 10, 40, 40));
        string expected = bound == "encoded-bytes" ? "per-image limit" : bound == "decoded-pixels" ? "pixel limit" : "publication image byte limit";
        Assert.Contains(expected, failure.Message);
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport); Assert.Empty(document.Images);
        PublisherShape shape = document.AddShape(pages[0].Id, OfficeShape.Rectangle(10, 10), 10, 10);
        Assert.Equal(2U, shape.Id);
    }

    [Fact]
    public void Repeated_master_image_projection_failure_can_retry_on_a_physical_page_without_consuming_either_id() {
        byte[] png = Png();
        var options = new PublisherReadOptions { MaximumTotalImageBytes = png.Length * 4 };
        PublisherDocument document = PublisherDocument.Create(300, 400, options);
        PublisherPage master = document.AddMasterPage("Header");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        var pages = document.Pages; var masters = document.MasterPages; var report = document.ReadReport;
        Assert.Throws<InvalidDataException>(() => document.AddPicture(master.Id, png, 10, 10, 40, 40));
        Assert.Same(pages, document.Pages); Assert.Same(masters, document.MasterPages); Assert.Same(report, document.ReadReport);
        Assert.Empty(document.Images); Assert.Equal(0, document.ReadReport.CurrentObjectCount);
        PublisherPicture added = document.AddPicture(pages[0].Id, png, 10, 10, 40, 40);
        Assert.Equal(4U, added.Id); Assert.Equal(1, added.ImageId);
        Assert.Empty(document.MasterPages[0].Pictures); Assert.Single(document.Pages[0].Pictures); Assert.Empty(document.Pages[1].Pictures);
        Assert.Equal(1, document.ReadReport.CurrentObjectCount); Assert.Equal(0, document.SourceReadReport.ImageCount);
    }

    [Fact]
    public void Picture_copy_edits_retain_object_limits_and_failed_addition_does_not_reserve_an_asset() {
        PublisherDocument document = PublisherDocument.Create(300, 400, new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 2 } });
        uint page = document.Pages[0].Id;
        document.AddTextFrame(page, 0, 0, 40, 20, Array.Empty<OfficeRichTextParagraph>());
        PublisherPicture picture = document.AddPicture(page, Png(), 10, 50, 40, 40);
        document.SetPictureBounds(picture.Id, 20, 50, 40, 40);
        Assert.Equal(2, document.ReadReport.CurrentObjectCount);
        var pages = document.Pages; var images = document.Images;
        Assert.Throws<InvalidDataException>(() => document.AddPicture(page, Png(), 70, 50, 40, 40));
        Assert.Throws<InvalidDataException>(() => document.AddPicture(page, picture.ImageId!.Value, 70, 50, 40, 40));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images);
        document.ReplacePictureImage(picture.Id, PublisherPictureEditingFixture.Image(OfficeColor.White));
        Assert.Equal(2, document.ReadReport.CurrentObjectCount); Assert.Equal(2, Assert.Single(document.Images).Id);
        PublisherPage added = document.AddPage(); Assert.Equal(picture.Id + 1, added.Id);
    }

    [Fact]
    public void Exhausted_native_image_reservation_does_not_consume_the_next_object_identity() {
        const string fixture = "NativeComparison/Image/native-image.pub";
        PublisherDocument original = PublisherDocument.Load(PublisherNativeTests.Fixture(fixture));
        uint nativeId = Assert.Single(original.Pages[0].Pictures).Id;
        byte[] input = PublisherDrawingFixture.Mutate(new Dictionary<ushort, uint> { [0x4104] = int.MaxValue }, 1, fixture: fixture, objectId: nativeId);
        PublisherDocument document = PublisherDocument.Load(input);
        var pages = document.Pages; var images = document.Images;
        Assert.Equal(int.MaxValue, Assert.Single(pages[0].Pictures).ImageId);
        Assert.Throws<InvalidOperationException>(() => document.AddPicture(pages[0].Id, Png(), 10, 10, 40, 40));
        Assert.Same(pages, document.Pages); Assert.Same(images, document.Images);
        PublisherShape shape = document.AddShape(pages[0].Id, OfficeShape.Rectangle(10, 10), 10, 10);
        PublisherShape expected = original.AddShape(original.Pages[0].Id, OfficeShape.Rectangle(10, 10), 10, 10);
        Assert.Equal(expected.Id, shape.Id);
        PublisherDocument exhausted = PublisherShapeEditingFixture.Create(exhaustedIds: true);
        Assert.Throws<InvalidOperationException>(() => exhausted.AddPicture(exhausted.Pages[0].Id, Png(), 10, 10, 40, 40));
    }

    private static byte[] Png() => PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(20, 130, 40));
    private static OfficeDrawingImage Painted(PublisherPage page, uint id) => Assert.Single(PublisherNativeTests.Elements(page.CreateDrawing())
        .OfType<OfficeDrawingImage>(), image => image.SourceElementIds?.Contains("publisher-object-" + id) == true);
    private static byte[] OrientedJpeg() => OfficeJpegCodec.Encode(new OfficeRasterImage(2, 3, OfficeColor.White), new OfficeJpegEncodeOptions {
        Metadata = new OfficeJpegMetadata(exif: new byte[] {
            0x49, 0x49, 0x2A, 0, 8, 0, 0, 0, 1, 0, 0x12, 1, 3, 0, 1, 0, 0, 0, 6, 0, 0, 0, 0, 0, 0, 0
        })
    });
    // Complete 2x2 RGBA PNG with two independently compressed APNG frames.
    private static byte[] AnimatedPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAACGFjVEwAAAACAAAAAPONk3AAAAAaZmNUTAAAAAAAAAACAAAAAgAAAAAAAAAAAAEACgAA6FTcAAAAABFJREFUeJxj+M/A8B+EGWAMAEfKB/lnWW63AAAAGmZjVEwAAAABAAAAAgAAAAIAAAAAAAAAAAABAAoAAHMnNtQAAAASZmRBVAAAAAJ4nGNg+A+FMAYAQ84H+d5oJ9kAAAAASUVORK5CYII=");
}
