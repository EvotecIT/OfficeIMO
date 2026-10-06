using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    // SVG positions are local to the text's transform. A negative local baseline or an
    // off-canvas local x can become visible after translation, rotation or scaling.
    // Retain the complete run in a bounded local surface; the destination SVG viewport
    // applies its clip after all of those transforms, as it does for other geometry.
    private static void AddTransformedTextRun(OfficeDrawing drawing, SvgTextRun run,
        OfficeTransform transform, SvgElementReferenceRegistry references,
        double maximumDimension, double maximumPixels, ref int unsupported) {
        double width = run.Width / run.GlyphScale;
        double height = run.FontSize * 1.25D;
        var font = new OfficeFontInfo(run.Style.FontFamily, run.FontSize, run.Style.FontFace, run.Style.FontStyle);
        var measure = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: drawing.Fonts,
            cancellationToken: references.CancellationToken);
        using var faceScope = measure.PushTextFace(font.Face);
        double size = Math.Max(1D, run.FontSize);
        double advance = run.HasExplicitAdvance ? width :
            measure.MeasurePositionedText(run.Text, size, font.FamilyName, font.Style, OfficeTextFeatureSettings.Default, run.TextDirection);
        var bounds = measure.MeasurePositionedTextBounds(run.Text, 0D, 0D, width, height, size,
            font, advance, OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", size,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, run.TextDirection);
        // An advance is not the ink extent: italic bearings, simulated styles and
        // fallback glyphs can paint beyond it. Retain that paint in the local surface.
        double left = Math.Floor(bounds.Left), top = Math.Floor(bounds.Top);
        double surfaceWidth = Math.Ceiling(bounds.Right) - left;
        double surfaceHeight = Math.Ceiling(bounds.Bottom) - top;
        if (!IsSupportedSvgViewport(surfaceWidth, surfaceHeight, maximumDimension, maximumPixels)
            || !references.TryChargeIntermediateSurface(surfaceWidth, surfaceHeight)) {
            unsupported++;
            return;
        }
        var local = new OfficeDrawing(surfaceWidth, surfaceHeight);
        local.Fonts.AddRange(drawing.Fonts);
        OfficeColor fill = run.Style.Fill!.Value;
        double opacity = Math.Max(0D, Math.Min(1D, run.Style.FillOpacity * run.Style.Opacity));
        OfficeColor color = OfficeColor.FromRgba(fill.R, fill.G, fill.B, (byte)Math.Round(fill.A * opacity));
        try {
            local.AddPositionedTextWithResolvedDirection(run.Text, -left, -top, width, height,
                font, color, height, run.HasExplicitAdvance ? width : (double?)null, run.TextDirection);
            drawing.AddEffectDrawing(local, OfficeTransform.Translate(left, top).Then(OfficeTransform.Scale(run.GlyphScale, 1D))
                .Then(OfficeTransform.Translate(run.X, run.Baseline - run.FontSize)).Then(transform));
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
        }
    }
}
