using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static void ApplyDrawingProfile(OfficeDrawing drawing, OfficeRenderingProfile? profile) {
        if (profile == null) return;
        drawing.Fonts.AddRange(profile.FontsSnapshot);
        drawing.Fonts.FontProgramProvider = profile.FontsSnapshot.FontProgramProvider;
        drawing.Fonts.FontVariationResolver = profile.FontsSnapshot.FontVariationResolver;
        drawing.TextShapingProvider = profile.TextShapingProvider;
        drawing.TextShapingLanguage = profile.TextShapingLanguage;
    }

    private static void CopyDrawingResources(OfficeDrawing source, OfficeDrawing target) {
        target.Fonts.AddRange(source.Fonts);
        target.Fonts.FontProgramProvider = source.Fonts.FontProgramProvider;
        target.Fonts.FontVariationResolver = source.Fonts.FontVariationResolver;
        target.TextShapingProvider = source.TextShapingProvider;
        target.TextShapingLanguage = source.TextShapingLanguage;
    }

    // Width and ink bounds must use the same selected font and shaping context as the final renderer.
    private static OfficeDrawingTextMetrics ResolveTextMetrics(OfficeDrawing drawing, OfficeDrawingTextMetrics? authoritative,
        CancellationToken cancellationToken) {
        if (authoritative != null) return authoritative;
        OfficeRasterCanvas canvas = OfficeDrawingTextLayout.CreateMetrics(drawing, cancellationToken);
        return new OfficeDrawingTextMetrics(canvas.MeasureText, canvas.MeasureTextPaintBounds,
            (text, size, family, style) => canvas.MeasureTextLineHorizontalPaintBounds(text ?? string.Empty, size, family, style));
    }
}
