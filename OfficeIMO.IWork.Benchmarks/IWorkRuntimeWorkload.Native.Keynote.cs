using OfficeIMO.PowerPoint;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.IWork.Benchmarks;

public sealed partial class IWorkRuntimeWorkload {
    // The text and backgrounds also occur in the independently exported Apple PPTX.
    private static readonly string[][] NativeKeynoteText = [
        ["hello keynote", "first bullet"],
        ["wrapped title keeps all of its lines in this fixed frame"]
    ];
    private static readonly string[] NativeKeynoteColors = ["56C1FF", "FF968D"];

    private void ValidateNativeKeynote(IWorkKeynoteProjection projection) {
        if (!projection.HasEditableContent || projection.Slides.Count != _units
            || projection.SlideSize is not { WidthPoints: 1024, HeightPoints: 768 })
            throw new InvalidDataException("Native Keynote projection is incomplete or its canvas differs.");
        for (int i = 0; i < projection.Slides.Count; i++) {
            IWorkKeynoteSlide slide = projection.Slides[i];
            int profile = i % 2;
            string[] texts = slide.Drawables.Where(drawable => drawable.Kind == IWorkKeynoteDrawableKind.TextBox)
                .Select(drawable => drawable.TextBox!.Content.PlainText).ToArray();
            if (slide.Index != i + 1 || slide.IsSkipped || slide.Images.Count != 0 || slide.Tables.Count != 0
                || slide.Drawables.Count != texts.Length || slide.PresenterNotes != NativeKeynoteNotes(profile)
                || slide.BackgroundColor?.RgbHex != NativeKeynoteColors[profile]
                || !texts.SequenceEqual(NativeKeynoteText[profile]))
                throw new InvalidDataException($"Native Keynote slide {i + 1} differs: background {slide.BackgroundColor?.RgbHex}, text "
                    + System.Text.Json.JsonSerializer.Serialize(texts));
            for (int box = 0; box < slide.Drawables.Count; box++) {
                IWorkGeometry geometry = slide.Drawables[box].TextBox!.Geometry
                    ?? throw new InvalidDataException("Native Keynote frame is missing.");
                ValidateNativeKeynoteGeometry(box, geometry.LeftPoints, geometry.TopPoints, geometry.WidthPoints, geometry.HeightPoints);
                if (geometry.RotationDegrees != 0) throw new InvalidDataException("Native Keynote frame rotation differs.");
            }
            VerifiedUnits++;
        }
        if (_nativeKeynoteScale) {
            // Independent schema decode pins the complete presentation order despite repeated text.
            string identities = string.Join("\n", projection.Slides.Select(slide =>
                (slide.SourceIdentity ?? throw new InvalidDataException("Missing native slide identity."))
                    .RecordIdentifier.ToString(CultureInfo.InvariantCulture)));
            if (Convert.ToHexString(SHA256.HashData(Encoding.ASCII.GetBytes(identities)))
                != "3EBFE306987FE7DD805358ED9C37EFA42A056D45F6A4E61B894D2DEF151BC66A")
                throw new InvalidDataException("Native Keynote slide identity order differs from the independent decode.");
        }
    }

    private void ValidateNativeKeynoteOutput(Stream saved) {
        using PowerPointDocument presentation = PowerPointDocument.Load(saved);
        if (presentation.Slides.Count != _units || presentation.SlideSize.WidthPoints != 1024 || presentation.SlideSize.HeightPoints != 768
            || presentation.ValidateDocument().Count != 0)
            throw new InvalidDataException("Native PPTX structure differs.");
        for (int i = 0; i < presentation.Slides.Count; i++) {
            var slide = presentation.Slides[i];
            int profile = i % 2;
            PowerPointTextBox[] boxes = slide.TextBoxes.ToArray();
            if (slide.Hidden || slide.Shapes.Count != boxes.Length || slide.BackgroundColor != NativeKeynoteColors[profile]
                || slide.GetSpeakerNotesText() != NativeKeynoteNotes(profile)
                || !boxes.Select(box => box.Text.TrimEnd('\n')).SequenceEqual(NativeKeynoteText[profile]))
                throw new InvalidDataException("Native PPTX text, notes, order or background differs.");
            for (int box = 0; box < boxes.Length; box++) {
                var geometry = boxes[box];
                ValidateNativeKeynoteGeometry(box, geometry.Left.ToPoints(), geometry.Top.ToPoints(), geometry.Width.ToPoints(), geometry.Height.ToPoints());
                if (geometry.Rotation.GetValueOrDefault() != 0 || geometry.HorizontalFlip.GetValueOrDefault()
                    || geometry.VerticalFlip.GetValueOrDefault())
                    throw new InvalidDataException("Native PPTX frame rotation or flip differs.");
            }
            VerifiedUnits++;
        }
    }

    private static string NativeKeynoteNotes(int profile) => profile == 0 ? string.Empty : "note text here";

    private static void ValidateNativeKeynoteGeometry(int box, double left, double top, double width, double height) {
        double expectedTop = box == 0 ? 146 : 402;
        double expectedWidth = box == 0 ? 914.0989379882812 : 914;
        double expectedHeight = box == 0 ? 260 : 114.67706298828125;
        if (Math.Abs(left - 55) >= .001 || Math.Abs(top - expectedTop) >= .001
            || Math.Abs(width - expectedWidth) >= .001 || Math.Abs(height - expectedHeight) >= .001)
            throw new InvalidDataException("Native Keynote frame geometry differs from the independent Apple reference.");
    }
}
