using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;

internal static partial class NativeKeynoteWriterSmoke {
    private const string ImageAcceptedSha256 = "5523DBE029B9A508D2BF3E3AD1A54C6266AC6C3BDDF8F23D6051FCC5F1E2F765";

    private static void RunImages() {
        IWorkKeynoteDocument document = CreateImageExample();
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == ImageAcceptedSha256,
            "NativeAOT image graph differs from the independently decoded and Apple-opened package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "NativeAOT image graph is nondeterministic.");
        IWorkSourceDocument source = IWorkSourceDocument.Open(bytes);
        IWorkKeynoteProjection projection = source.ReadKeynote();
        Require(projection.Slides.Count == 3 && projection.Slides.All(slide => slide.Images.Count == 1),
            "NativeAOT image readback is incomplete.");
        Require(!source.Records.Any(record => record.Identifier is 71 or 72),
            "NativeAOT graph uses Apple's reserved metadata identities.");
        Require(projection.Slides[0].Drawables.Select(drawable => drawable.Kind).SequenceEqual(new[] {
            IWorkKeynoteDrawableKind.TextBox, IWorkKeynoteDrawableKind.TextBox,
            IWorkKeynoteDrawableKind.Image, IWorkKeynoteDrawableKind.TextBox
        }), "NativeAOT text/image insertion order differs.");
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(bytes),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var saved = PowerPointPresentation.Load(output);
        Require(saved.Slides.Count == 3 && !saved.ValidateDocument().Any(), "NativeAOT image PPTX did not save/reopen correctly.");
        for (int i = 0; i < 3; i++) {
            PowerPointPicture picture = saved.Slides[i].Pictures.Single();
            IWorkImageAsset image = projection.Slides[i].Images.Single();
            Require(picture.GetImageBytes().SequenceEqual(image.GetBytes()) && picture.AltText == image.AccessibilityDescription,
                "NativeAOT saved image lost its payload or accessibility description.");
        }
        using var unchanged = new MemoryStream(new byte[] { 9, 8, 7 });
        try {
            document.Save(unchanged, new IWorkKeynoteWriteOptions { MaximumImages = 2 });
            throw new InvalidOperationException("NativeAOT image-count budget was not enforced.");
        } catch (InvalidDataException) { }
        Require(unchanged.ToArray().SequenceEqual(new byte[] { 9, 8, 7 }), "NativeAOT image-budget failure changed the destination.");
        Console.WriteLine($"PASS | native Keynote images | {ImageAcceptedSha256} | {bytes.Length} bytes | PNG alpha/JPEG, reserved metadata identities, stacking, deterministic creation, saved/reopened PPTX and image limits");
    }

    private static IWorkKeynoteDocument CreateImageExample() {
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
}
