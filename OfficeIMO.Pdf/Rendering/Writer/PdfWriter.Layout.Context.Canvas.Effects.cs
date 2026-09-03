using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasEffect(PdfCanvasEffectItem item) {
            if (item.Opacity >= 1D && item.BlendMode == OfficeBlendMode.Normal) {
                OfficeTransform transform = ConvertTopLeftCanvasTransform(item.Transform, currentOpts.PageHeight);
                RenderOpaqueEffectGroupInline(transform, () => RenderCanvasBlock(new PdfCanvasBlock(item.Items)));
                return;
            }

            RenderEffectGroup(item.Transform, item.Opacity, item.BlendMode, () => RenderCanvasBlock(new PdfCanvasBlock(item.Items)));
        }

        private void RenderEffectGroup(OfficeTransform topLeftPageTransform, double opacity, Action renderContent) =>
            RenderEffectGroup(topLeftPageTransform, opacity, OfficeBlendMode.Normal, renderContent);

        private void RenderEffectGroup(OfficeTransform topLeftPageTransform, double opacity, OfficeBlendMode blendMode, Action renderContent) {
            OfficeTransform transform = ConvertTopLeftCanvasTransform(topLeftPageTransform, currentOpts.PageHeight);
            if (opacity >= 1D && blendMode == OfficeBlendMode.Normal &&
                (currentOpts.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers ||
                 _suppressCanvasAccessibilityWrappers ||
                 _suppressCanvasActualTextChildren)) {
                RenderOpaqueEffectGroupInline(transform, renderContent);
                return;
            }

            bool artifactContent = _suppressCanvasActualTextChildren;
            int annotationStart = currentPage!.Annotations.Count;
            int textAnnotationStart = currentPage.TextAnnotations.Count;
            int freeTextAnnotationStart = currentPage.FreeTextAnnotations.Count;
            int highlightAnnotationStart = currentPage.HighlightAnnotations.Count;
            int imageStart = currentPage.Images.Count;
            int formFieldStart = currentPage.FormFields.Count;
            string? opacityState = EnsureGraphicsState(opacity, opacity, blendMode);
            int contentStart = sb.Length;
            _canvasClipDepth++;
            try {
                renderContent();
            } finally {
                _canvasClipDepth--;
            }

            string groupContent = sb.ToString(contentStart, sb.Length - contentStart);
            sb.Length = contentStart;
            if (artifactContent) groupContent = "/Artifact BMC\n" + groupContent + "EMC\n";
            ResolveEffectGroupBounds(transform, out double boundsLeft, out double boundsBottom, out double boundsRight, out double boundsTop);
            string token = "\n%OIMO_EFFECT_GROUP_" + (currentPage.EffectGroups.Count + 1).ToString("D6", CultureInfo.InvariantCulture) + "\n";
            currentPage.EffectGroups.Add(new PageEffectGroup {
                Content = pageContents.Store(groupContent),
                Token = token,
                Transform = transform,
                BoundsLeft = boundsLeft,
                BoundsBottom = boundsBottom,
                BoundsRight = boundsRight,
                BoundsTop = boundsTop,
                GraphicsStateName = opacityState
            });
            sb.Append(token);
            TransformCanvasRectangles(currentPage.Annotations, annotationStart, transform);
            TransformCanvasRectangles(currentPage.TextAnnotations, textAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.FreeTextAnnotations, freeTextAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.HighlightAnnotations, highlightAnnotationStart, transform);
            TransformCanvasPageImageBounds(currentPage.Images, imageStart, transform);
            TransformCanvasRectangles(currentPage.FormFields, formFieldStart, transform);
            pageDirty = true;
        }

        private void RenderOpaqueEffectGroupInline(OfficeTransform transform, Action renderContent) {
            bool artifactContent = _suppressCanvasActualTextChildren;
            int annotationStart = currentPage!.Annotations.Count;
            int textAnnotationStart = currentPage.TextAnnotations.Count;
            int freeTextAnnotationStart = currentPage.FreeTextAnnotations.Count;
            int highlightAnnotationStart = currentPage.HighlightAnnotations.Count;
            int imageStart = currentPage.Images.Count;
            int formFieldStart = currentPage.FormFields.Count;
            if (artifactContent) {
                sb.Append("/Artifact BMC\n");
            }
            var content = new ContentStreamBuilder(sb).SaveState();
            if (!transform.Equals(OfficeTransform.Identity)) {
                content.TransformMatrix(transform);
            }
            _canvasClipDepth++;
            try {
                renderContent();
            } finally {
                _canvasClipDepth--;
                new ContentStreamBuilder(sb).RestoreState();
                if (artifactContent) {
                    sb.Append("EMC\n");
                }
            }

            TransformCanvasRectangles(currentPage.Annotations, annotationStart, transform);
            TransformCanvasRectangles(currentPage.TextAnnotations, textAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.FreeTextAnnotations, freeTextAnnotationStart, transform);
            TransformCanvasRectangles(currentPage.HighlightAnnotations, highlightAnnotationStart, transform);
            TransformCanvasPageImageBounds(currentPage.Images, imageStart, transform);
            TransformCanvasRectangles(currentPage.FormFields, formFieldStart, transform);
            pageDirty = true;
        }

        private void ResolveEffectGroupBounds(
            OfficeTransform transform,
            out double left,
            out double bottom,
            out double right,
            out double top) {
            left = 0D;
            bottom = 0D;
            right = currentOpts.PageWidth;
            top = currentOpts.PageHeight;
            if (!transform.TryInvert(out OfficeTransform inverse)) return;

            (left, bottom, right, top) = inverse.TransformRectangleBounds(
                0D,
                0D,
                currentOpts.PageWidth,
                currentOpts.PageHeight);
        }

        private static void TransformCanvasPageImageBounds(System.Collections.Generic.List<PageImage> images, int startIndex, OfficeTransform transform) {
            for (int index = startIndex; index < images.Count; index++) {
                PageImage image = images[index];
                (double x, double y, double width, double height) = EffectivePageImageBounds(image);
                (double left, double bottom, double right, double top) = TransformRectangle(x, y, x + width, y + height, transform);
                image.EffectiveX = left;
                image.EffectiveY = bottom;
                image.EffectiveW = right - left;
                image.EffectiveH = top - bottom;
            }
        }

        private static (double X, double Y, double Width, double Height) EffectivePageImageBounds(PageImage image) {
            if (image.EffectiveX.HasValue && image.EffectiveY.HasValue && image.EffectiveW.HasValue && image.EffectiveH.HasValue) {
                return (image.EffectiveX.Value, image.EffectiveY.Value, image.EffectiveW.Value, image.EffectiveH.Value);
            }

            OfficeTransform imageTransform = new OfficeImageProjection(
                new OfficeImagePlacement(image.X, image.Y, image.W, image.H),
                rotationDegrees: image.RotationAngle,
                rotationCenterX: image.RotationCenterX,
                rotationCenterY: image.RotationCenterY,
                flipHorizontal: image.HorizontalFlip,
                flipVertical: image.VerticalFlip)
                .CreateUnitSquareTransform();
            (double left, double bottom, double right, double top) = imageTransform.TransformRectangleBounds(0D, 0D, 1D, 1D);
            return (left, bottom, right - left, top - bottom);
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
