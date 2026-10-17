using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    /// <summary>Retains navigation metadata when an empty paint clip culls a subtree.</summary>
    private static void AddClippedNamedDestinations(
        PdfCore.PdfPageCanvas canvas,
        IEnumerable<HtmlRenderVisual> visuals,
        OfficeTransform transform,
        CancellationToken cancellationToken) {
        foreach (HtmlRenderVisual visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderNamedDestination destination) {
                OfficePoint point = transform.TransformPoint(new OfficePoint(destination.X, destination.Y));
                canvas.NamedDestination(MapNamedDestination(destination.Name),
                    point.X * PointsPerCssPixel, point.Y * PointsPerCssPixel);
                continue;
            }
            if (visual is HtmlRenderEffectGroup effect) {
                AddClippedNamedDestinations(canvas, effect.Visuals, effect.Transform.Then(transform), cancellationToken);
                continue;
            }
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup path => path.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderLayoutRegion layout => layout.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            if (children != null) AddClippedNamedDestinations(canvas, children, transform, cancellationToken);
        }
    }
}
