using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasEffect(PdfCanvasEffectItem item) {
            RenderEffectGroup(item.Transform, item.Opacity, () => RenderCanvasBlock(new PdfCanvasBlock(item.Items)));
        }

        private void RenderEffectGroup(OfficeTransform topLeftPageTransform, double opacity, Action renderContent) {
            OfficeTransform transform = ConvertTopLeftCanvasTransform(topLeftPageTransform, currentOpts.PageHeight);
            int annotationStart = currentPage!.Annotations.Count;
            int textAnnotationStart = currentPage.TextAnnotations.Count;
            int freeTextAnnotationStart = currentPage.FreeTextAnnotations.Count;
            int highlightAnnotationStart = currentPage.HighlightAnnotations.Count;
            int formFieldStart = currentPage.FormFields.Count;
            string? opacityState = EnsureGraphicsState(opacity, opacity);
            int contentStart = sb.Length;
            _canvasClipDepth++;
            try {
                renderContent();
            } finally {
                _canvasClipDepth--;
            }

            string groupContent = sb.ToString(contentStart, sb.Length - contentStart);
            sb.Length = contentStart;
            string token = "\n%OIMO_EFFECT_GROUP_" + (currentPage.EffectGroups.Count + 1).ToString("D6", CultureInfo.InvariantCulture) + "\n";
            currentPage.EffectGroups.Add(new PageEffectGroup {
                Content = pageContents.Store(groupContent),
                Token = token,
                Transform = transform,
                GraphicsStateName = opacityState
            });
            sb.Append(token);
            TransformCanvasRectangles(currentPage.Annotations, annotationStart, transform);
            TransformCanvasRectangles(currentPage.TextAnnotations, textAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.FreeTextAnnotations, freeTextAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.HighlightAnnotations, highlightAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.FormFields, formFieldStart, transform);
            pageDirty = true;
        }

        private static OfficeTransform ConvertTopLeftCanvasTransform(OfficeTransform transform, double pageHeight) =>
            new OfficeTransform(
                transform.M11,
                -transform.M12,
                -transform.M21,
                transform.M22,
                transform.M21 * pageHeight + transform.OffsetX,
                pageHeight * (1D - transform.M22) - transform.OffsetY);
    }
}
