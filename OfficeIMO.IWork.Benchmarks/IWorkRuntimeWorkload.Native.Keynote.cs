using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Benchmarks;

public sealed partial class IWorkRuntimeWorkload {
    // The text and backgrounds also occur in the independently exported Apple PPTX.
    private static readonly string[][] NativeKeynoteText = [
        ["hello keynote", "first bullet"],
        ["wrapped title keeps all of its lines in this fixed frame"]
    ];
    private static readonly string[] NativeKeynoteColors = ["56C1FF", "FF968D"];

    private void ValidateNativeKeynote(IWorkKeynoteProjection projection) {
        if (!projection.HasEditableContent || projection.Slides.Count != 2
            || projection.SlideSize is not { WidthPoints: 1024, HeightPoints: 768 })
            throw new InvalidDataException("Native Keynote projection is incomplete or its canvas differs.");
        for (int i = 0; i < projection.Slides.Count; i++) {
            IWorkKeynoteSlide slide = projection.Slides[i];
            string[] texts = slide.Drawables.Where(drawable => drawable.Kind == IWorkKeynoteDrawableKind.TextBox)
                .Select(drawable => drawable.TextBox!.Content.PlainText).ToArray();
            if (slide.BackgroundColor?.RgbHex != NativeKeynoteColors[i]
                || !texts.SequenceEqual(NativeKeynoteText[i]))
                throw new InvalidDataException($"Native Keynote slide {i + 1} differs: background {slide.BackgroundColor?.RgbHex}, text "
                    + System.Text.Json.JsonSerializer.Serialize(texts));
            VerifiedUnits++;
        }
    }

    private void ValidateNativeKeynoteOutput(Stream saved) {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(saved);
        if (presentation.Slides.Count != 2 || presentation.SlideSize.WidthPoints != 1024 || presentation.SlideSize.HeightPoints != 768
            || presentation.ValidateDocument().Count != 0)
            throw new InvalidDataException("Native PPTX structure differs.");
        for (int i = 0; i < presentation.Slides.Count; i++) {
            var slide = presentation.Slides[i];
            if (slide.BackgroundColor != NativeKeynoteColors[i]
                || !slide.TextBoxes.Select(box => box.Text.TrimEnd('\n')).SequenceEqual(NativeKeynoteText[i]))
                throw new InvalidDataException("Native PPTX text, order or background differs.");
            VerifiedUnits++;
        }
    }
}
