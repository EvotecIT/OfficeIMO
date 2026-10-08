using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    // Formatting frames describe flow. Resolve ink through each child's effects
    // and authored clips before expanding an intermediate Drawing buffer.
    private static (double Left, double Top, double Right, double Bottom) ResolveDrawingBufferBounds(
        IEnumerable<HtmlRenderVisual> visuals, double surfaceWidth, double surfaceHeight,
        OfficeFontFaceCollection fonts, CancellationToken cancellationToken,
        HtmlRenderVisual? authoredClip = null) {
        double left = Math.Min(0D, MinimumLeft(visuals));
        double top = Math.Min(0D, MinimumTop(visuals));
        double right = Math.Max(surfaceWidth, MaximumRight(visuals));
        double bottom = Math.Max(surfaceHeight, MaximumBottom(visuals));
        OfficeRasterCanvas? measurement = null;
        foreach (HtmlRenderVisual visual in visuals) {
            var ink = MeasureInk(visual);
            if (ink.HasValue && authoredClip != null) ink = ClipInk(ink.Value, authoredClip);
            if (!ink.HasValue) continue;
            // Integer support preserves antialiased edges and the raster origin.
            if (ink.Value.Left < left) left = Math.Floor(ink.Value.Left) - 1D;
            if (ink.Value.Top < top) top = Math.Floor(ink.Value.Top) - 1D;
            if (ink.Value.Right > right) right = Math.Ceiling(ink.Value.Right) + 1D;
            if (ink.Value.Bottom > bottom) bottom = Math.Ceiling(ink.Value.Bottom) + 1D;
        }
        return (left, top, right, bottom);

        (double Left, double Top, double Right, double Bottom)? MeasureInk(HtmlRenderVisual visual) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderText original && original.Text.Length > 0 && original.TextAdvanceWidth is double measuredAdvance) {
                measurement ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: fonts, cancellationToken: cancellationToken);
                HtmlRenderText text = original.ResolveBaselineForPainting();
                string value = text.BidiVisualOrderResolved ? "\u202D" + text.Text + "\u202C" : text.Text;
                double advance = text.TextPaintWidth ?? (measuredAdvance > 0D ? measuredAdvance : text.Width);
                var ink = measurement.MeasurePositionedTextBounds(value, text.X, text.Y,
                    text.Width, text.Height, text.Font.Size, text.Font, advance, text.Alignment,
                    text.FeatureSettings, text.FontPalette, original.Font.Size, text.UnderlineStyle, text.StrikethroughStyle);
                return ink.HasInk ? (ink.Left, ink.Top, ink.Right, ink.Bottom) : null;
            }
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup path => path.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            (double Left, double Top, double Right, double Bottom)? result = null;
            if (children == null) return result;
            foreach (HtmlRenderVisual child in children) {
                var ink = MeasureInk(child);
                if (!ink.HasValue) continue;
                result = !result.HasValue ? ink : (Math.Min(result.Value.Left, ink.Value.Left),
                    Math.Min(result.Value.Top, ink.Value.Top), Math.Max(result.Value.Right, ink.Value.Right),
                    Math.Max(result.Value.Bottom, ink.Value.Bottom));
            }
            if (!result.HasValue) return result;
            if (visual is HtmlRenderEffectGroup transformed) {
                var ink = result.Value;
                OfficePoint p1 = transformed.Transform.TransformPoint(new OfficePoint(ink.Left, ink.Top));
                OfficePoint p2 = transformed.Transform.TransformPoint(new OfficePoint(ink.Right, ink.Top));
                OfficePoint p3 = transformed.Transform.TransformPoint(new OfficePoint(ink.Left, ink.Bottom));
                OfficePoint p4 = transformed.Transform.TransformPoint(new OfficePoint(ink.Right, ink.Bottom));
                return (Math.Min(Math.Min(p1.X, p2.X), Math.Min(p3.X, p4.X)),
                    Math.Min(Math.Min(p1.Y, p2.Y), Math.Min(p3.Y, p4.Y)),
                    Math.Max(Math.Max(p1.X, p2.X), Math.Max(p3.X, p4.X)),
                    Math.Max(Math.Max(p1.Y, p2.Y), Math.Max(p3.Y, p4.Y)));
            }
            return ClipInk(result.Value, visual);
        }
    }

    private static (double Left, double Top, double Right, double Bottom)? ClipInk(
        (double Left, double Top, double Right, double Bottom) ink, HtmlRenderVisual clip) {
        if (clip is HtmlRenderClipGroup rectangle) {
            if (rectangle.ClipHorizontal) { ink.Left = Math.Max(ink.Left, rectangle.ClipX); ink.Right = Math.Min(ink.Right, rectangle.ClipX + rectangle.ClipWidth); }
            if (rectangle.ClipVertical) { ink.Top = Math.Max(ink.Top, rectangle.ClipY); ink.Bottom = Math.Min(ink.Bottom, rectangle.ClipY + rectangle.ClipHeight); }
        } else if (clip is HtmlRenderPathClipGroup path) {
            ink.Left = Math.Max(ink.Left, path.ClipX); ink.Right = Math.Min(ink.Right, path.ClipX + path.ClipPath.Width);
            ink.Top = Math.Max(ink.Top, path.ClipY); ink.Bottom = Math.Min(ink.Bottom, path.ClipY + path.ClipPath.Height);
        }
        return ink.Right > ink.Left && ink.Bottom > ink.Top ? ink : null;
    }
}
