using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeImageWritingTests {
    [Fact]
    public void Native_resaved_picture_without_a_filename_reference_can_be_replaced_again() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/NormalizedPictureNames/native-normalized-picture-names.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherPicture first = document.Pages[0].Pictures.Single(), last = document.Pages[2].Pictures.Single();
        Assert.Equal(1, first.ImageId); Assert.Equal(3, last.ImageId);
        Assert.True(last.CanEdit, last.EditLimitation);
        byte[] original = document.Images.Single(image => image.Id == first.ImageId).GetBytes();
        byte[] jpeg = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/ImagePagesJpegReplacement/native-imagepagesjpegreplacement-replacement.jpg"));
        document.ReplacePictureImage(last.Id, jpeg);
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(output);
        PublisherPicture unchanged = reopened.Pages[0].Pictures.Single(), replaced = reopened.Pages[2].Pictures.Single();
        Assert.Equal((first.Id, first.ImageId, first.X, first.Y, first.Width, first.Height),
            (unchanged.Id, unchanged.ImageId, unchanged.X, unchanged.Y, unchanged.Width, unchanged.Height));
        Assert.Equal((last.Id, last.PageId, last.X, last.Y, last.Width, last.Height, last.Crop),
            (replaced.Id, replaced.PageId, replaced.X, replaced.Y, replaced.Width, replaced.Height, replaced.Crop));
        Assert.Equal(4, replaced.ImageId);
        Assert.Equal(jpeg, reopened.Images.Single(image => image.Id == replaced.ImageId).GetBytes());
        Assert.Equal(original, reopened.Images.Single(image => image.Id == 1).GetBytes());
        Assert.Equal(original, reopened.Images.Single(image => image.Id == 3).GetBytes());
        var before = Compound(source); var after = Compound(output);
        // Content edits invalidate the separately qualified cached thumbnail.
        foreach (string name in before.Streams.Keys.Where(name => name != "Contents" && !name.StartsWith("Escher/", StringComparison.Ordinal)
            && name != OfficeOlePropertySetWriter.SummaryInformationStreamName))
            Assert.Equal(before.Streams[name], after.Streams[name]);
        // The newly owned filename resource must also support the next save.
        reopened.ReplacePictureImage(replaced.Id, original);
        PublisherDocument repeated = PublisherDocument.Load(reopened.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(original, repeated.Images.Single(image => image.Id == repeated.Pages[2].Pictures.Single().ImageId).GetBytes());
        Assert.Equal(jpeg, repeated.Images.Single(image => image.Id == 4).GetBytes());
        Assert.Equal(first.ImageId, repeated.Pages[0].Pictures.Single().ImageId);
    }

    [Fact]
    public void Independent_native_picture_replacement_retains_the_duplicate_and_master_on_the_original_asset() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/PictureReplacement/native-picturereplacement.pub"));
        PublisherPicture replaced = document.Pages[0].Pictures.Single(picture => picture.Id == 298);
        PublisherPicture duplicate = document.Pages[0].Pictures.Single(picture => picture.Id == 300);
        PublisherPicture master = document.MasterPages.SelectMany(page => page.Pictures).Single();
        Assert.Equal(2, document.Images.Count); Assert.Equal(2, replaced.ImageId);
        Assert.Equal(1, duplicate.ImageId); Assert.Equal(1, master.ImageId);
        Assert.Equal((72D, 144D, 96D, 96D), (replaced.X, replaced.Y, replaced.Width, replaced.Height));
        Assert.Equal(File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/PictureReplacement/native-picturereplacement-replacement.png")),
            document.Images.Single(image => image.Id == replaced.ImageId).GetBytes());
        Assert.Equal(File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/PictureReplacement/native-picturereplacement.png")),
            document.Images.Single(image => image.Id == duplicate.ImageId).GetBytes());
    }

    [Theory]
    [InlineData("Image", false)]
    [InlineData("PictureOwnership", false)]
    [InlineData("PictureOwnership", true)]
    [InlineData("Newsletter", true)]
    [InlineData("PictureCropped", false)]
    [InlineData("NormalizedSharedPictures", true)]
    public void Native_image_saves_replace_one_placement_and_retain_source_assets_including_after_repeated_edits(string fixture, bool repeat) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/native-{fixture.ToLowerInvariant()}.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherPicture[] original = Pictures(document);
        PublisherPicture target = document.Pages[0].Pictures[0];
        var assets = document.Images.ToDictionary(image => image.Id, image => image.GetBytes());
        byte[] replacement = repeat ? OfficeJpegCodec.Encode(new OfficeRasterImage(48, 24, OfficeColor.Blue), new OfficeJpegEncodeOptions())
            : PublisherPictureEditingFixture.Image(OfficeColor.Blue, 48, 24);
        if (repeat) document.ReplacePictureImage(target.Id, PublisherPictureEditingFixture.Image(OfficeColor.Red, 12, 16));
        document.ReplacePictureImage(target.Id, replacement);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        byte[] candidate = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(candidate);
        Assert.Equal(original.Select(picture => picture.Id), Pictures(reopened).Select(picture => picture.Id));
        foreach (PublisherPicture before in original) {
            PublisherPicture after = Pictures(reopened).Single(picture => picture.Id == before.Id);
            Assert.Equal((before.PageId, before.X, before.Y, before.Width, before.Height, before.Crop),
                (after.PageId, after.X, after.Y, after.Width, after.Height, after.Crop));
            byte[] expected = before.Id == target.Id ? replacement : assets[before.ImageId!.Value];
            Assert.Equal(expected, reopened.Images.Single(image => image.Id == after.ImageId).GetBytes());
            if (before.Id != target.Id) Assert.Equal(before.ImageId, after.ImageId);
        }
        Assert.Equal(assets.Count + 1, reopened.Images.Count);
        foreach (var asset in assets) Assert.Equal(asset.Value, reopened.Images.Single(image => image.Id == asset.Key).GetBytes());
        Assert.Equal(document.TextStories.Select(story => (story.Id, story.Text)), reopened.TextStories.Select(story => (story.Id, story.Text)));
        Assert.Equal(document.Pages.Select(page => (page.Id, page.MasterPageId)), reopened.Pages.Select(page => (page.Id, page.MasterPageId)));
        var beforeStorage = Compound(source); var afterStorage = Compound(candidate);
        foreach (string name in beforeStorage.Streams.Keys.Where(name => name != "Contents" && !name.StartsWith("Escher/", StringComparison.Ordinal) && name != "\u0005SummaryInformation"))
            Assert.Equal(beforeStorage.Streams[name], afterStorage.Streams[name]);
    }

    [Fact]
    public void Native_picture_and_frame_growth_share_one_source_offset_space_without_losing_bounds_or_new_assets() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Newsletter/native-newsletter.pub"));
        PublisherPicture picture = document.Pages.SelectMany(page => page.Pictures).First();
        PublisherTextFrame frame = document.Pages.SelectMany(page => page.TextFrames).First(item => item.CanEditLayout);
        byte[] replacement = PublisherPictureEditingFixture.Image(OfficeColor.Blue, 48, 24);
        document.ReplacePictureImage(picture.Id, replacement);
        document.SetPictureBounds(picture.Id, picture.X + 3, picture.Y + 6, picture.Width - 6, picture.Height - 12);
        document.SetTextFrameLayout(frame.Id, 2, 12.1, new OfficeTextPadding(1.21, 2.42, 3.63, 4.84), OfficeTextVerticalAlignment.Center);
        byte[] candidate = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(candidate);
        PublisherPicture edited = Pictures(reopened).Single(item => item.Id == picture.Id);
        Assert.Equal((picture.X + 3, picture.Y + 6, picture.Width - 6, picture.Height - 12), (edited.X, edited.Y, edited.Width, edited.Height));
        Assert.Equal(replacement, reopened.Images.Single(image => image.Id == edited.ImageId).GetBytes());
        PublisherTextFrame layout = reopened.Pages.SelectMany(page => page.TextFrames).Single(item => item.Id == frame.Id);
        Assert.Equal(2, layout.ColumnCount); Assert.Equal(12.1, layout.ColumnSpacing);
        Assert.Equal(new OfficeTextPadding(1.21, 2.42, 3.63, 4.84), layout.Padding); Assert.Equal(OfficeTextVerticalAlignment.Center, layout.VerticalAlignment);
        Assert.Equal(document.TextStories.Select(story => story.Text), reopened.TextStories.Select(story => story.Text));
    }

    private static PublisherPicture[] Pictures(PublisherDocument document) => document.Pages.Concat(document.MasterPages).SelectMany(page => page.Pictures).ToArray();
    [Theory]
    [InlineData("filename-reference")]
    [InlineData("allocation-counter")]
    [InlineData("unresolved-sibling")]
    [InlineData("reference-count")]
    public void Native_image_saves_refuse_unqualified_metadata_and_source_slot_aliases_instead_of_emitting_partial_edits(string invalid) {
        const string fixture = "NativeComparison/PictureOwnership/native-pictureownership.pub";
        byte[] input = invalid == "unresolved-sibling"
            ? PublisherDrawingFixture.Mutate(new Dictionary<ushort, uint> { [0x4104] = 2 }, 1, fixture: fixture, objectId: 300)
            : PublisherInputContractTests.Mutate(invalid == "reference-count" ? "Escher/EscherStm" : "Contents", bytes => {
                var data = new PublisherBinaryData(bytes, "source");
                var context = new PublisherParseContext(new PublisherReadOptions(), default);
                if (invalid == "reference-count") {
                    var reader = new PublisherEscherReader(data, null, context);
                    PublisherEscherRecord group = reader.Records(0, bytes.Length).Single(record => record.Kind == 0xF000);
                    PublisherEscherRecord store = reader.Records(group.Offset, group.End).Single(record => record.Kind == 0xF001);
                    PublisherEscherRecord first = reader.Records(store.Offset, store.End)[0];
                    BitConverter.GetBytes(0U).CopyTo(bytes, first.Offset + 24);
                } else if (invalid == "allocation-counter") {
                    var fields = new PublisherBlockReader(data, context).Chunk(data.Offset(data.U32(0x1A)));
                    PublisherBlock count = fields.Single(field => field.Id == 1);
                    BitConverter.GetBytes(count.Value + 1).CopyTo(bytes, count.Offset);
                } else {
                    PublisherContents contents = new PublisherContentsReader(data, context).Read();
                    PublisherBlock reference = contents.Shapes[298].Fields.Single(field => field.Id == 0x3A);
                    bytes[reference.Offset - 1] = 0x20;
                }
            }, fixture);
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Equal(input, document.ToPubBytes());
        PublisherPicture picture = document.Pages[0].Pictures.Single(item => item.Id == 298);
        Assert.True(picture.CanEdit, picture.EditLimitation);
        document.ReplacePictureImage(picture.Id, PublisherPictureEditingFixture.Image(OfficeColor.Blue));
        PublisherNativePreparation preparation = document.PrepareNativeBytes(64 * 1024 * 1024);
        Assert.False(preparation.IsPrepared);
        Assert.Contains(preparation.FidelityDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure
            && diagnostic.Message.Contains(invalid is "filename-reference" or "allocation-counter" ? "metadata" : "image-store"));
    }

    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        return compound!;
    }
}
