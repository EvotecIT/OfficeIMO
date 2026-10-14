using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;

internal static class NativeImageResizeSmoke {
    internal static void Run() {
        var source = IWorkSourceDocument.Open(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "preservation.key")));
        var images = source.ReadKeynote().Slides.SelectMany(slide => slide.Images).ToArray();
        var edit = source.BeginEdit();
        for (int index = 0; index < 3; index++) {
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
        var options = new IWorkNativeSaveOptions { AllowUnassessedVisualCaches = true };
        byte[] bytes = edit.SaveBytes(options);
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == "49D1293B65A1AA6EAB6FD9A15FF0A1D482EF76760344F61A5D48F0D5E1F5A851"
            && bytes.SequenceEqual(edit.SaveBytes(options)), "Image resizing differs from the independently decoded Apple-accepted package.");
        Verify(source, bytes, edited: false);
        byte[] native = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "preservation-resized-edited.key"));
        Require(Convert.ToHexString(SHA256.HashData(native)) == "B079128AD659C0130E28BEA97D62824FF504B3A0BC34E77C6B380421210FE099",
            "Apple-edited resized source changed.");
        Verify(source, native, edited: true);
        try {
            edit.ResizeKeynoteImage(images[2].SourceIdentity!, 960, 168);
            throw new InvalidOperationException("Out-of-canvas resize was accepted.");
        } catch (ArgumentOutOfRangeException) { }
        Require(bytes.SequenceEqual(edit.SaveBytes(options)), "Rejected resize changed prior edits.");
        for (int index = 0; index < 3; index++) {
            var image = images[index];
            edit.MoveKeynoteImage(image.SourceIdentity!, (float)image.Geometry!.LeftPoints, (float)image.Geometry.TopPoints);
            edit.ResizeKeynoteImage(image.SourceIdentity!, (float)image.Geometry.WidthPoints, (float)image.Geometry.HeightPoints);
        }
        var restored = edit.SaveBytesResult();
        Require(restored.Report.UpdatedRecords.Count == 0
            && restored.Report.OutputPackageContentHash == source.ComputePackageContentHash(), "Restoring original image frames lost source bytes.");
        Console.WriteLine("PASS | native Keynote image resizing | exact composed package, PNG/alpha/JPEG resources, Apple edit/reopen, strict saved PPTX, pending-frame bounds and restoration");
    }

    private static void Verify(IWorkSourceDocument original, byte[] bytes, bool edited) {
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadKeynote(); var before = original.ReadKeynote();
        Require(projection.Diagnostics.Count == 0 && projection.SourceDeclarationIssues.Count == 0
            && projection.SourceReferenceIssues.Count == 0 && projection.Slides.Count == 3, "Resized source has incomplete declarations.");
        using var result = source.ToPowerPointPresentationResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = PowerPointPresentation.Load(output);
        Require(saved.Slides.Count == 3 && !saved.ValidateDocument().Any(), "Resized source failed saved PowerPoint reconstruction.");
        for (int index = 0; index < 3; index++) {
            var image = projection.Slides[index].Images.Single(); var expected = before.Slides[index].Images.Single();
            var frame = image.Geometry!; var picture = saved.Slides[index].Pictures.Single();
            Require(image.GetBytes().SequenceEqual(expected.GetBytes()) && image.AccessibilityDescription == expected.AccessibilityDescription
                && image.PixelWidth == 96 && image.PixelHeight == 64
                && before.Slides[index].Drawables.Select(item => item.Kind).SequenceEqual(projection.Slides[index].Drawables.Select(item => item.Kind))
                && frame.LeftPoints == (index == 2 ? 100 : 80) && frame.TopPoints == (index == 0 ? 125 : 135)
                && frame.WidthPoints == (index == 1 ? 360 : index == 2 && edited ? 421 : 420) && frame.HeightPoints == (index == 1 ? 180 : 168),
                "Resized source lost image pixels, description, stacking or final geometry.");
            Require(picture.GetImageBytes().SequenceEqual(image.GetBytes()) && picture.AltText == image.AccessibilityDescription
                && picture.Left == frame.LeftPoints * 12700 && picture.Top == frame.TopPoints * 12700
                && picture.Width == frame.WidthPoints * 12700 && picture.Height == frame.HeightPoints * 12700,
                "Saved PowerPoint lost resized image content or geometry.");
        }
    }

    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
