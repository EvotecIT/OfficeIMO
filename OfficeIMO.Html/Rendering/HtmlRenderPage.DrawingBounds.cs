using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    // A formatting frame describes flow, not all glyph ink. Measure with the
    // same scoped fonts and positioned paint path used by the Drawing owner
    // before an effect or clip rasterizes its child buffer.
    private static (double Left, double Top, double Right, double Bottom) ResolveDrawingBufferBounds(
        IEnumerable<HtmlRenderVisual> visuals, double surfaceWidth, double surfaceHeight,
        OfficeFontFaceCollection fonts, CancellationToken cancellationToken) {
        double left = Math.Min(0D, MinimumLeft(visuals));
        double top = Math.Min(0D, MinimumTop(visuals));
        double right = Math.Max(surfaceWidth, MaximumRight(visuals));
        double bottom = Math.Max(surfaceHeight, MaximumBottom(visuals));
        OfficeRasterCanvas? measurement = null;
        foreach (HtmlRenderVisual visual in visuals) Include(visual, OfficeTransform.Identity);
        return (left, top, right, bottom);

        void Include(HtmlRenderVisual visual, OfficeTransform transform) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderText original && original.Text.Length > 0 && original.TextAdvanceWidth is double measuredAdvance) {
                measurement ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: fonts, cancellationToken: cancellationToken);
                HtmlRenderText text = original.ResolveBaselineForPainting();
                string value = text.BidiVisualOrderResolved ? "\u202D" + text.Text + "\u202C" : text.Text;
                double advance = text.TextPaintWidth ?? (measuredAdvance > 0D ? measuredAdvance : text.Width);
                var ink = measurement.MeasurePositionedTextBounds(value, text.X, text.Y,
                    text.Width, text.Height, text.Font.Size, text.Font, advance, text.Alignment,
                    text.FeatureSettings, text.FontPalette, original.Font.Size, text.UnderlineStyle, text.StrikethroughStyle);
                if (!ink.HasInk) return;
                OfficePoint p1 = transform.TransformPoint(new OfficePoint(ink.Left, ink.Top));
                OfficePoint p2 = transform.TransformPoint(new OfficePoint(ink.Right, ink.Top));
                OfficePoint p3 = transform.TransformPoint(new OfficePoint(ink.Left, ink.Bottom));
                OfficePoint p4 = transform.TransformPoint(new OfficePoint(ink.Right, ink.Bottom));
                var projected = (Left: Math.Min(Math.Min(p1.X, p2.X), Math.Min(p3.X, p4.X)),
                    Top: Math.Min(Math.Min(p1.Y, p2.Y), Math.Min(p3.Y, p4.Y)),
                    Right: Math.Max(Math.Max(p1.X, p2.X), Math.Max(p3.X, p4.X)),
                    Bottom: Math.Max(Math.Max(p1.Y, p2.Y), Math.Max(p3.Y, p4.Y)));
                // Integer sampling support avoids clipping an antialiased edge
                // and keeps the raster origin aligned when restoring local space.
                if (projected.Left < left) left = Math.Floor(projected.Left) - 1D;
                if (projected.Top < top) top = Math.Floor(projected.Top) - 1D;
                if (projected.Right > right) right = Math.Ceiling(projected.Right) + 1D;
                if (projected.Bottom > bottom) bottom = Math.Ceiling(projected.Bottom) + 1D;
                return;
            }
            if (visual is HtmlRenderEffectGroup effect) {
                foreach (HtmlRenderVisual child in effect.Visuals) Include(child, effect.Transform.Then(transform));
            } else if (visual is HtmlRenderClipGroup clip) {
                foreach (HtmlRenderVisual child in clip.Visuals) Include(child, transform);
            } else if (visual is HtmlRenderPathClipGroup path) {
                foreach (HtmlRenderVisual child in path.Visuals) Include(child, transform);
            } else if (visual is HtmlRenderSemanticGroup semantic) {
                foreach (HtmlRenderVisual child in semantic.Visuals) Include(child, transform);
            } else if (visual is HtmlRenderLayoutRegion region) {
                foreach (HtmlRenderVisual child in region.Visuals) Include(child, transform);
            } else if (visual is HtmlRenderLogicalTextGroup logical) {
                foreach (HtmlRenderVisual child in logical.Visuals) Include(child, transform);
            }
        }
    }
}
