using System.Threading;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkKeynoteImageResizeTests {
    [Fact]
    public void Composed_image_edits_match_the_independently_decoded_Apple_accepted_package() {
        var source = Source(); var edit = source.BeginEdit();
        var images = source.ReadKeynote().Slides.SelectMany(slide => slide.Images).ToArray();
        for (int index = 0; index < images.Length; index++) {
            var image = images[index]; float width = index == 1 ? 360 : 420, height = index == 1 ? 180 : 168;
            float left = (float)image.Geometry!.LeftPoints + 20, top = (float)image.Geometry.TopPoints + 15;
            if (index == 1) {
                edit.ResizeKeynoteImage(image.SourceIdentity!, width, height);
                edit.MoveKeynoteImage(image.SourceIdentity!, left, top);
            } else {
                edit.MoveKeynoteImage(image.SourceIdentity!, left, top);
                edit.ResizeKeynoteImage(image.SourceIdentity!, width, height);
            }
        }
        var saved = edit.SaveBytesResult(AcceptCaches());
        Assert.Equal(File.ReadAllBytes(Fixture("keynote-preserved-resized-v15.4.key")), saved.GetPackageBytes());
        Assert.Equal(new ulong[] { 125, 139, 154 }, saved.Report.UpdatedRecords.Select(record => record.RecordIdentifier).Order());
        Assert.Equal(3, saved.Report.UpdatedEntryPaths.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Resized_and_Apple_edited_sources_retain_images_frames_stacking_and_strict_saved_PowerPoint(bool edited) {
        var original = Source().ReadKeynote();
        string name = "keynote-preserved-resized" + (edited ? "-edited" : "") + "-v15.4.key";
        var source = IWorkSourceDocument.Open(Fixture(name)); var projection = source.ReadKeynote();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        using var conversion = source.ToPowerPointDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        conversion.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); conversion.Value.Save(output); output.Position = 0;
        using var saved = PowerPointDocument.Load(output);
        using var native = PowerPointDocument.Load(Fixture("keynote-preserved-resized-edited-v15.4.pptx"));
        Assert.Empty(saved.ValidateDocument()); Assert.Equal(3, saved.Slides.Count);
        for (int index = 0; index < 3; index++) {
            var expected = Assert.Single(original.Slides[index].Images);
            var image = Assert.Single(projection.Slides[index].Images);
            Assert.Equal(expected.GetBytes(), image.GetBytes()); Assert.Equal(expected.AccessibilityDescription, image.AccessibilityDescription);
            Assert.Equal(expected.SourceIdentity!.RecordIdentifier, image.SourceIdentity!.RecordIdentifier);
            Assert.Equal(96, image.PixelWidth); Assert.Equal(64, image.PixelHeight);
            Assert.Equal(index == 2 ? 100 : 80, image.Geometry!.LeftPoints);
            Assert.Equal(index == 0 ? 125 : 135, image.Geometry.TopPoints);
            Assert.Equal(index == 1 ? 360 : index == 2 && edited ? 421 : 420, image.Geometry.WidthPoints);
            Assert.Equal(index == 1 ? 180 : 168, image.Geometry.HeightPoints);
            Assert.Equal(original.Slides[index].Drawables.Select(drawable => drawable.Kind),
                projection.Slides[index].Drawables.Select(drawable => drawable.Kind));
            var picture = Assert.Single(saved.Slides[index].Pictures); var oracle = Assert.Single(native.Slides[index].Pictures);
            Assert.Equal(image.GetBytes(), picture.GetImageBytes()); Assert.Equal(oracle.GetImageBytes(), picture.GetImageBytes());
            Assert.Equal(image.AccessibilityDescription, picture.AltText); Assert.Equal(oracle.AltText, picture.AltText);
            Assert.Equal(OfficeLength.Points(image.Geometry.LeftPoints), picture.Left); Assert.Equal(OfficeLength.Points(image.Geometry.TopPoints), picture.Top);
            Assert.Equal(OfficeLength.Points(image.Geometry.WidthPoints), picture.Width); Assert.Equal(OfficeLength.Points(image.Geometry.HeightPoints), picture.Height);
            Assert.Equal(oracle.Left, picture.Left); Assert.Equal(oracle.Top, picture.Top); Assert.Equal(oracle.Height, picture.Height);
            Assert.Equal(oracle.Width - OfficeLength.Points(index == 2 && !edited ? 1 : 0), picture.Width);
            Assert.Equal(native.Slides[index].Shapes.Select(shape => shape.GetType()), saved.Slides[index].Shapes.Select(shape => shape.GetType()));
        }
    }

    [Fact]
    public void Move_and_resize_compose_validate_the_pending_frame_and_restore_original_bytes() {
        var source = Source(); var projection = source.ReadKeynote(); var image = Jpeg(source);
        var first = source.BeginEdit(); var second = source.BeginEdit();
        first.MoveKeynoteImage(image.SourceIdentity!, 100, 130);
        first.ResizeKeynoteImage(image.SourceIdentity!, 300, 180);
        second.ResizeKeynoteImage(image.SourceIdentity!, 300, 180);
        second.MoveKeynoteImage(image.SourceIdentity!, 100, 130);
        var saved = first.SaveBytesResult(AcceptCaches());
        Assert.Equal(saved.GetPackageBytes(), second.SaveBytes(AcceptCaches()));
        var edited = Jpeg(IWorkSourceDocument.Open(saved.GetPackageBytes()));
        Assert.Equal(100, edited.Geometry!.LeftPoints); Assert.Equal(130, edited.Geometry.TopPoints);
        Assert.Equal(300, edited.Geometry.WidthPoints); Assert.Equal(180, edited.Geometry.HeightPoints);
        Assert.Equal(image.GetBytes(), edited.GetBytes()); Assert.Single(saved.Report.UpdatedRecords);
        Assert.Equal(600, image.Geometry!.WidthPoints); Assert.Equal(240, image.Geometry.HeightPoints);

        first.MoveKeynoteImage(image.SourceIdentity!, (float)projection.SlideSize!.WidthPoints - 300,
            (float)projection.SlideSize.HeightPoints - 180);
        byte[] beforeRejection = first.SaveBytes(AcceptCaches());
        Assert.Throws<ArgumentOutOfRangeException>(() => first.ResizeKeynoteImage(image.SourceIdentity!, 600, 240));
        Assert.Equal(beforeRejection, first.SaveBytes(AcceptCaches()));
        first.MoveKeynoteImage(image.SourceIdentity!, (float)image.Geometry.LeftPoints, (float)image.Geometry.TopPoints);
        first.ResizeKeynoteImage(image.SourceIdentity!, (float)image.Geometry.WidthPoints, (float)image.Geometry.HeightPoints);
        var restored = first.SaveBytesResult();
        Assert.Empty(restored.Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), restored.Report.OutputPackageContentHash);
    }

    [Fact]
    public void Invalid_or_cancelled_resize_preserves_prior_edits_and_requires_cache_acceptance() {
        var source = Source(); var image = Jpeg(source); var session = source.BeginEdit();
        session.MoveKeynoteImage(image.SourceIdentity!, 100, 130);
        byte[] before = session.SaveBytes(AcceptCaches());
        foreach (float invalid in new[] { 0, -1, float.NaN, float.NegativeInfinity, float.PositiveInfinity, float.MaxValue }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => session.ResizeKeynoteImage(image.SourceIdentity!, invalid, 120));
            Assert.Throws<ArgumentOutOfRangeException>(() => session.ResizeKeynoteImage(image.SourceIdentity!, 120, invalid));
        }
        Assert.Throws<ArgumentException>(() => session.ResizeKeynoteImage(source.ReadKeynote().Slides[0].SourceIdentity!, 120, 120));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => session.ResizeKeynoteImage(image.SourceIdentity!, 120, 120, cancellation.Token));
        Assert.Equal(before, session.SaveBytes(AcceptCaches()));
        session.ResizeKeynoteImage(image.SourceIdentity!, 300, 180);
        using var output = new MemoryStream(); output.WriteByte(9);
        Assert.Throws<InvalidOperationException>(() => session.Save(output));
        Assert.Equal(new byte[] { 9 }, output.ToArray());
    }

    private static IWorkSourceDocument Source() => IWorkSourceDocument.Open(Fixture("keynote-created-images-edited-v15.4.key"));
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static IWorkImageAsset Jpeg(IWorkSourceDocument source) => source.ReadKeynote().Slides.SelectMany(slide => slide.Images)
        .Single(image => image.MediaType == "image/jpeg");
    private static IWorkNativeSaveOptions AcceptCaches() => new() { AllowUnassessedVisualCaches = true };
}
