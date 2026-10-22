using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Reports paint removed by a running-content clip, ignoring unused formatting-frame overflow.</summary>
    private void ReportPageMarginContentClip(IReadOnlyList<HtmlRenderVisual> visuals,
        double x, double y, double width, double height, bool clipHorizontal, bool clipVertical, string source) {
        if (_reportedMarginTextOverflow.Contains(source)) return;
        foreach (HtmlRenderVisual visual in visuals) {
            foreach (HtmlRenderRectangle paint in EnumeratePageMarginPaintBounds(visual)) {
                bool horizontal = clipHorizontal && (paint.X < x - 0.0001D || paint.Right > x + width + 0.0001D);
                bool vertical = clipVertical && (paint.Y < y - 0.0001D || paint.Bottom > y + height + 0.0001D);
                if (!horizontal && !vertical) continue;
                _reportedMarginTextOverflow.Add(source);
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                    "Generated running-element paint exceeded its page-margin content area and was clipped.",
                    HtmlDiagnosticSeverity.Warning, source, horizontal ? "margin-width" : "margin-height", OfficeConversionLossKind.Omission);
                return;
            }
        }
    }

    /// <summary>Checks generated paint rather than an authored margin box's potentially empty frame.</summary>
    private void ReportPageMarginPhysicalOverflow(IReadOnlyList<HtmlRenderVisual> visuals, HtmlRenderPage page, string source) {
        foreach (HtmlRenderVisual visual in visuals) {
            foreach (HtmlRenderRectangle paint in EnumeratePageMarginPaintBounds(visual)) {
                if (paint.X >= -0.0001D && paint.Y >= -0.0001D
                    && paint.Right <= page.Width + 0.0001D && paint.Bottom <= page.Height + 0.0001D) continue;
                if (_reportedMarginTextOverflow.Add(source + " page-paint")) {
                    _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                        "Generated page-margin paint extended beyond the physical page and was clipped.",
                        HtmlDiagnosticSeverity.Warning, source, "margin-page-bounds", OfficeConversionLossKind.Omission);
                }
                return;
            }
        }
    }

    private IEnumerable<HtmlRenderRectangle> EnumeratePageMarginPaintBounds(HtmlRenderVisual visual) {
        CheckCancellation();
        ChargeLayoutOperation("page-margin physical paint");
        if (visual is HtmlRenderEffectGroup { Opacity: <= 0D }) yield break;
        if (visual.PaintChildren is { } children) {
            foreach (HtmlRenderVisual child in children) {
                foreach (HtmlRenderRectangle childPaint in EnumeratePageMarginPaintBounds(child)) {
                    HtmlRenderRectangle? paint = ClipPageMarginPaintBounds(childPaint, visual);
                    if (paint.HasValue) yield return visual is HtmlRenderEffectGroup effect
                        ? HtmlRenderRectangle.Transform(effect.Transform, paint.Value) : paint.Value;
                }
            }
        } else if (visual is HtmlRenderText text) {
            if (text.Color.A == 0 || string.IsNullOrWhiteSpace(text.Text)) yield break;
            // The same measured fragments feed both PDF and Drawing painting;
            // a wide aligned formatting frame is not itself lost paint.
            foreach (HtmlRenderText fragment in text.GetWrappedPaintFragments()) {
                CheckCancellation();
                ChargeLayoutOperation("page-margin measured line paint");
                if (string.IsNullOrWhiteSpace(fragment.Text)) continue;
                HtmlRenderText painted = fragment.ResolveBaselineForPainting();
                double advance = painted.TextAdvanceWidth is double measured && measured > 0D ? measured : painted.Width;
                yield return new HtmlRenderRectangle(painted.X, painted.Y - painted.PaintTopOverflow,
                    painted.TextPaintWidth ?? advance,
                    painted.Height + painted.PaintTopOverflow);
            }
        } else if (visual is HtmlRenderShape shape) {
            OfficeShape paint = shape.InnerShape;
            bool fill = (paint.FillOpacity ?? 1D) > 0D && (paint.FillRadialGradient != null
                ? paint.FillRadialGradient.Stops.Any(stop => stop.Color.A > 0)
                : paint.FillGradient != null ? paint.FillGradient.Stops.Any(stop => stop.Color.A > 0)
                : paint.FillColor?.A > 0);
            bool stroke = paint.StrokeWidth > 0D && (paint.StrokeOpacity ?? 1D) > 0D && (paint.StrokeRadialGradient != null
                ? paint.StrokeRadialGradient.Stops.Any(stop => stop.Color.A > 0)
                : paint.StrokeGradient != null ? paint.StrokeGradient.Stops.Any(stop => stop.Color.A > 0)
                : paint.StrokeColor?.A > 0);
            if (fill || stroke) yield return new HtmlRenderRectangle(visual.X, visual.Y, visual.Width, visual.Height);
        } else if (visual is HtmlRenderImage || visual is HtmlRenderDrawing drawing && drawing.InnerDrawing.Elements.Count > 0) {
            yield return new HtmlRenderRectangle(visual.X, visual.Y, visual.Width, visual.Height);
        }
    }

    private static HtmlRenderRectangle? ClipPageMarginPaintBounds(HtmlRenderRectangle paint, HtmlRenderVisual visual) {
        double left = paint.X, top = paint.Y, right = paint.Right, bottom = paint.Bottom;
        if (visual is HtmlRenderClipGroup clip) {
            if (clip.ClipHorizontal) { left = Math.Max(left, clip.ClipX); right = Math.Min(right, clip.ClipX + clip.ClipWidth); }
            if (clip.ClipVertical) { top = Math.Max(top, clip.ClipY); bottom = Math.Min(bottom, clip.ClipY + clip.ClipHeight); }
        } else if (visual is HtmlRenderPathClipGroup path) {
            left = Math.Max(left, path.ClipX); right = Math.Min(right, path.ClipX + path.ClipPath.Width);
            top = Math.Max(top, path.ClipY); bottom = Math.Min(bottom, path.ClipY + path.ClipPath.Height);
        }
        return right > left && bottom > top ? new HtmlRenderRectangle(left, top, right - left, bottom - top) : null;
    }
}
