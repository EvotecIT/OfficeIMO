using System.IO.Compression;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkKeynoteImageWritingTests {
    [Fact]
    public void Image_creation_matches_the_Apple_qualified_graph_and_exported_PowerPoint() {
        byte[] expected = File.ReadAllBytes(Fixture("keynote-created-images-v15.4.key"));
        Assert.Equal(expected, ImageExample().SaveBytes());
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(expected),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var saved = PowerPointPresentation.Load(output);
        using var apple = PowerPointPresentation.Load(Fixture("keynote-created-images-v15.4.pptx"));
        Assert.Equal(3, saved.Slides.Count); Assert.Empty(saved.ValidateDocument());
        for (int i = 0; i < 3; i++) {
            Assert.Equal(apple.Slides[i].Shapes.Select(shape => shape.GetType()),
                saved.Slides[i].Shapes.Select(shape => shape.GetType()));
            PowerPointPicture actual = Assert.Single(saved.Slides[i].Pictures);
            PowerPointPicture oracle = Assert.Single(apple.Slides[i].Pictures);
            Assert.Equal(oracle.GetImageBytes(), actual.GetImageBytes());
            Assert.Equal(oracle.AltText, actual.AltText);
            Assert.Equal(oracle.Left, actual.Left); Assert.Equal(oracle.Top, actual.Top);
            Assert.Equal(oracle.Width, actual.Width); Assert.Equal(oracle.Height, actual.Height);
        }
    }

    [Fact]
    public void Apple_saved_image_edit_keeps_pixels_stacking_and_description_after_reopen() {
        IWorkKeynoteProjection original = IWorkSourceDocument.Open(Fixture("keynote-created-images-v15.4.key")).ReadKeynote();
        string editedPath = Fixture("keynote-created-images-edited-v15.4.key");
        IWorkKeynoteProjection edited = IWorkSourceDocument.Open(editedPath).ReadKeynote();
        Assert.Equal(3, edited.Slides.Count);
        for (int i = 0; i < 3; i++) {
            Assert.Equal(original.Slides[i].Drawables.Select(drawable => drawable.Kind),
                edited.Slides[i].Drawables.Select(drawable => drawable.Kind));
            IWorkImageAsset expected = Assert.Single(original.Slides[i].Images);
            IWorkImageAsset actual = Assert.Single(edited.Slides[i].Images);
            Assert.Equal(expected.GetBytes(), actual.GetBytes());
            Assert.Equal(expected.AccessibilityDescription, actual.AccessibilityDescription);
            Assert.Equal(i == 2 ? 80 : 60, actual.Geometry!.LeftPoints);
            Assert.Equal(expected.Geometry!.TopPoints, actual.Geometry.TopPoints);
            Assert.Equal(expected.Geometry.WidthPoints, actual.Geometry.WidthPoints);
            Assert.Equal(expected.Geometry.HeightPoints, actual.Geometry.HeightPoints);
        }
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(editedPath,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var reopened = PowerPointPresentation.Load(output);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Equal(80 * 12700, reopened.Slides[2].Pictures.Single().Left);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_images_keep_payload_geometry_description_and_mixed_stacking_in_saved_PowerPoint(bool jpeg) {
        byte[] imageBytes = ImageBytes(jpeg);
        var document = IWorkKeynoteDocument.Create(320, 240);
        IWorkKeynoteSlideBuilder slide = document.AddSlide();
        slide.AddText("BEHIND", 10, 10, 280, 60);
        IWorkKeynoteImage added = slide.AddImage(imageBytes, 20, 40, 120, 80, "Native image 😀");
        slide.AddText("IN FRONT", 30, 50, 240, 60);
        slide.AddImage(imageBytes, 160, 40, 80, 120);
        byte[] package = document.SaveBytes();
        Assert.Equal(package, document.SaveBytes());
        Assert.Equal(jpeg ? "image/jpeg" : "image/png", added.MediaType);
        Assert.Equal(24, added.PixelWidth); Assert.Equal(16, added.PixelHeight);
        IWorkSourceDocument source = IWorkSourceDocument.Open(new MemoryStream(package));
        IWorkWireMessage metadata = IWorkProtobuf.Parse(Assert.Single(source.Records,
            record => record.MessageType == 11006).GetPayload(), new IWorkReadOptions());
        IWorkWireMessage[] resources = metadata.GetRepeatedMessages(4).ToArray();
        Assert.Equal(2, resources.Length);
        foreach (IWorkWireMessage resource in resources) {
            // An empty DataAttributes record passed managed reading but Keynote
            // rejected the package: native image data requires its class marker.
            IWorkWireMessage attributes = resource.GetMessage(10)!;
            Assert.NotNull(attributes);
            IWorkWireMessage imageAttributes = attributes.GetMessage(100)!;
            Assert.NotNull(imageAttributes);
            Assert.Equal(24, imageAttributes.GetMessage(1)!.GetFloat(1));
            Assert.Equal(16, imageAttributes.GetMessage(1)!.GetFloat(2));
            Assert.Equal(0UL, imageAttributes.GetUnsigned(3));
        }
        IWorkKeynoteProjection projection = source.ReadKeynote();
        Assert.Empty(projection.SourceDeclarationIssues);
        IWorkKeynoteSlide sourceSlide = Assert.Single(projection.Slides);
        Assert.Equal(new[] { IWorkKeynoteDrawableKind.TextBox, IWorkKeynoteDrawableKind.Image,
            IWorkKeynoteDrawableKind.TextBox, IWorkKeynoteDrawableKind.Image }, sourceSlide.Drawables.Select(d => d.Kind));
        IWorkImageAsset image = sourceSlide.Images[0];
        Assert.Equal(imageBytes, image.GetBytes());
        Assert.Equal("Native image 😀", image.AccessibilityDescription);
        Assert.Equal(20, image.Geometry!.LeftPoints); Assert.Equal(40, image.Geometry.TopPoints);
        Assert.Equal(120, image.Geometry.WidthPoints); Assert.Equal(80, image.Geometry.HeightPoints);
        Assert.False(image.HasMask);
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(package),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        PowerPointShape[] shapes = reopened.Slides[0].Shapes.ToArray();
        Assert.Equal(4, shapes.Length);
        Assert.IsType<PowerPointTextBox>(shapes[0]); Assert.IsType<PowerPointPicture>(shapes[1]);
        Assert.IsType<PowerPointTextBox>(shapes[2]); Assert.IsType<PowerPointPicture>(shapes[3]);
        Assert.Equal("Native image 😀", shapes[1].AltText);
        using var zip = new ZipArchive(new MemoryStream(saved.ToArray()));
        Assert.Contains(zip.Entries, entry => entry.FullName.StartsWith("ppt/media/", StringComparison.Ordinal)
            && Read(entry).SequenceEqual(imageBytes));
    }

    [Fact]
    public void Added_image_bytes_are_snapshots_and_invalid_images_do_not_mutate_the_slide() {
        var document = IWorkKeynoteDocument.Create(); IWorkKeynoteSlideBuilder slide = document.AddSlide();
        byte[] input = ImageBytes(false), expected = (byte[])input.Clone();
        IWorkKeynoteImage image = slide.AddImage(input, 0, 0, 100, 100);
        Array.Clear(input); Array.Clear(image.GetBytes());
        Assert.Equal(expected, image.GetBytes());
        Assert.Throws<ArgumentException>(() => slide.AddImage(expected.Take(expected.Length - 5).ToArray(), 0, 0, 100, 100));
        Assert.Throws<ArgumentException>(() => slide.AddImage(new byte[] { 71, 73, 70, 56, 57, 97 }, 0, 0, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddImage(expected, 950, 0, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddImage(expected, 0, 0, float.NaN, 100));
        Assert.Single(slide.Images);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkKeynoteImage>)slide.Images).Clear());
        Assert.Equal(expected, Assert.Single(IWorkSourceDocument.Open(new MemoryStream(document.SaveBytes()))
            .ReadKeynote().Slides[0].Images).GetBytes());
    }

    [Theory]
    [InlineData("count")]
    [InlineData("encoded")]
    [InlineData("decoded")]
    [InlineData("package")]
    public void Aggregate_image_limits_reject_before_stream_or_existing_file_mutation(string boundary) {
        byte[] payload = ImageBytes(false);
        var document = IWorkKeynoteDocument.Create(); var slide = document.AddSlide();
        slide.AddImage(payload, 0, 0, 100, 100); slide.AddImage(payload, 100, 0, 100, 100);
        var limits = new IWorkKeynoteWriteOptions();
        switch (boundary) {
            case "count": limits.MaximumImages = 1; break;
            case "encoded": limits.MaximumEncodedImageBytes = payload.Length; break;
            case "decoded": limits.MaximumDecodedImageBytes = 24 * 16 * 4; break;
            case "package": limits.MaximumPackageBytes = payload.Length; break;
        }
        byte[] sentinel = { 9, 8, 7 };
        using var stream = new MemoryStream(); stream.Write(sentinel);
        Assert.Throws<InvalidDataException>(() => document.Save(stream, limits)); Assert.Equal(sentinel, stream.ToArray());
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-KeynoteImages-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "existing.key"); File.WriteAllBytes(path, sentinel);
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: limits));
            Assert.Equal(sentinel, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
    }

    [Fact]
    public void Precancelled_image_validation_leaves_the_builder_and_destination_untouched() {
        var document = IWorkKeynoteDocument.Create(); var slide = document.AddSlide();
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => slide.AddImage(ImageBytes(false), 0, 0, 100, 100,
            cancellationToken: cancel.Token));
        Assert.Empty(slide.Images);
        slide.AddImage(ImageBytes(false), 0, 0, 100, 100);
        using var output = new MemoryStream();
        Assert.Throws<OperationCanceledException>(() => document.Save(output, cancellationToken: cancel.Token));
        Assert.Empty(output.ToArray());
    }

    [Fact]
    public void Native_identity_includes_image_payload_and_mixed_object_order() {
        IWorkKeynoteDocument Make(bool changed, bool front) {
            var document = IWorkKeynoteDocument.Create(); var slide = document.AddSlide();
            if (!front) slide.AddText("LABEL", 0, 0, 100, 100);
            slide.AddImage(ImageBytes(false, changed), 0, 0, 100, 100);
            if (front) slide.AddText("LABEL", 0, 0, 100, 100);
            return document;
        }
        byte[] original = Make(false, false).SaveBytes();
        Assert.Equal(original, Make(false, false).SaveBytes());
        string Identity(byte[] package) {
            using var zip = new ZipArchive(new MemoryStream(package));
            return System.Text.Encoding.ASCII.GetString(Read(zip.GetEntry("Metadata/DocumentIdentifier")!));
        }
        Assert.NotEqual(Identity(original), Identity(Make(true, false).SaveBytes()));
        Assert.NotEqual(Identity(original), Identity(Make(false, true).SaveBytes()));
    }

    internal static byte[] ImageBytes(bool jpeg, bool changed = false) {
        var raster = new OfficeRasterImage(24, 16, changed ? new OfficeColor(0, 255, 0) : new OfficeColor(255, 0, 0));
        return jpeg ? OfficeJpegCodec.Encode(raster) : OfficePngWriter.Encode(raster);
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "native-exports", name);

    private static IWorkKeynoteDocument ImageExample() {
        var raster = new OfficeRasterImage(96, 64);
        var alpha = new OfficeRasterImage(96, 64);
        for (int y = 0; y < 64; y++) for (int x = 0; x < 96; x++) {
            raster.SetPixel(x, y, (x < 48, y < 32) switch {
                (true, true) => new OfficeColor(220, 40, 50), (false, true) => new OfficeColor(20, 150, 80),
                (true, false) => new OfficeColor(20, 90, 210), _ => new OfficeColor(245, 180, 20)
            });
            alpha.SetPixel(x, y, new OfficeColor(160, 30, 210, (byte)(x < 32 ? 0 : x < 64 ? 128 : 255)));
        }
        var document = IWorkKeynoteDocument.Create();
        var first = document.AddSlide("F3F7FC");
        first.AddText("Native PNG — image behind text", 60, 30, 840, 60, "Helvetica Neue", 32, "003366");
        first.AddText("Covered source label", 100, 200, 500, 60, "Arial", 28);
        first.AddImage(OfficePngWriter.Encode(raster), 60, 110, 480, 320, "Opaque four-color PNG 😀");
        first.AddText("TEXT IN FRONT", 100, 220, 400, 60, "Arial", 32, "FFFFFF");
        var second = document.AddSlide("003366");
        second.AddText("Native alpha PNG", 60, 30, 840, 60, "Helvetica Neue", 32, "FFFFFF");
        second.AddImage(OfficePngWriter.Encode(alpha), 60, 120, 480, 320, "Transparent, half-alpha, opaque purple");
        second.AddRichText(new[] { new IWorkKeynoteTextRun("0% / 50% / ", "Arial", 28, "FFFFFF"),
            new IWorkKeynoteTextRun("100% alpha", "Arial", 28, "FFD700", bold: true) }, 60, 450, 840, 60);
        var third = document.AddSlide();
        third.AddText("Native JPEG", 60, 30, 840, 60, "Helvetica Neue", 32, "003366");
        third.AddImage(OfficeJpegCodec.Encode(raster), 60, 120, 600, 240, "Four-color JPEG, intentionally stretched frame");
        third.AddText("JPEG / stretched frame", 60, 400, 840, 60, "Times New Roman", 28);
        return document;
    }

    private static byte[] Read(ZipArchiveEntry entry) {
        using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
    }
}
