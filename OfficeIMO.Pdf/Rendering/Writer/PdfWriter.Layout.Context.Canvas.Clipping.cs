using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasClip(PdfCanvasClipItem item) {
            if (item.ClipPath.Kind == OfficeIMO.Drawing.OfficeClipPathKind.Empty) {
                if (double.IsNaN(item.X) || double.IsNaN(item.Y) || double.IsInfinity(item.X) || double.IsInfinity(item.Y)) {
                    throw new ArgumentOutOfRangeException(nameof(item), "Canvas clip coordinates must be finite.");
                }
            } else {
                ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas clip");
            }
            double bottomY = currentOpts.PageHeight - item.Y - item.Height;
            int annotationStart = currentPage!.Annotations.Count;
            int textAnnotationStart = currentPage.TextAnnotations.Count;
            int freeTextAnnotationStart = currentPage.FreeTextAnnotations.Count;
            int highlightAnnotationStart = currentPage.HighlightAnnotations.Count;
            int imageStart = currentPage.Images.Count;
            int formFieldStart = currentPage.FormFields.Count;
            new ContentStreamBuilder(sb).SaveState();
            AppendClipPath(sb, item.ClipPath, item.X, bottomY, item.Height);

            int diagnosticClipCount = _diagnosticCanvasClips.Count;
            if (currentOpts.HasDiagnosticsReport) {
                OfficeTransform placement = new OfficeTransform(1D, 0D, 0D, -1D, item.X, bottomY + item.Height).Then(_canvasEffectToPage);
                var diagnosticContours = OfficeClipPathGeometry.CreateContours(item.ClipPath, points => {
                    var mapped = new List<OfficePoint>(points.Count);
                    foreach (OfficePoint point in points) { cancellationToken.ThrowIfCancellationRequested(); mapped.Add(placement.TransformPoint(point)); }
                    return mapped;
                });
                _diagnosticCanvasClips.Add(OfficeTextInkClip.FromFilledContours(diagnosticContours, item.ClipPath.FillRule));
                _diagnosticCanvasSourceClips.Add(item.PreservesSourceClip);
            }
            _canvasClipDepth++;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _canvasClipDepth--;
                _diagnosticCanvasClips.RemoveRange(diagnosticClipCount, _diagnosticCanvasClips.Count - diagnosticClipCount);
                _diagnosticCanvasSourceClips.RemoveRange(diagnosticClipCount, _diagnosticCanvasSourceClips.Count - diagnosticClipCount);
            }

            ClipCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, item.Width, item.Height, item.ClipPath);
            ClipCanvasTextAnnotations(currentPage.TextAnnotations, textAnnotationStart, item.X, bottomY, item.Width, item.Height, item.ClipPath);
            ClipCanvasFreeTextAnnotations(currentPage.FreeTextAnnotations, freeTextAnnotationStart, item.X, bottomY, item.Width, item.Height, item.ClipPath);
            ClipCanvasHighlightAnnotations(currentPage.HighlightAnnotations, highlightAnnotationStart, item.X, bottomY, item.Width, item.Height, item.ClipPath);
            ClipCanvasPageImages(currentPage.Images, imageStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasFormFields(currentPage.FormFields, formFieldStart, item.X, bottomY, item.Width, item.Height, item.ClipPath.Kind);
            new ContentStreamBuilder(sb)
                .RestoreState();
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
            pageDirty = true;
        }

        private void ClipCanvasPageImages(System.Collections.Generic.List<PageImage> images, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = images.Count - 1; i >= startIndex; i--) {
                PageImage image = images[i];
                (double effectiveX, double effectiveY, double effectiveW, double effectiveH) = EffectivePageImageBounds(image);
                double x1 = System.Math.Max(effectiveX, clipX);
                double y1 = System.Math.Max(effectiveY, clipBottomY);
                double x2 = System.Math.Min(effectiveX + effectiveW, clipRight);
                double y2 = System.Math.Min(effectiveY + effectiveH, clipTop);
                if (x2 <= x1 || y2 <= y1) {
                    if (!string.IsNullOrEmpty(image.InlineDrawToken)) {
                        sb.Replace(image.InlineDrawToken, string.Empty);
                    }

                    ReleasePendingStructureDestination(image.NavigationTarget);
                    images.RemoveAt(i);
                    continue;
                }

                if (!string.IsNullOrEmpty(image.InlineDrawToken) || CanvasClipContainsFrame(image, x1, y1, x2, y2)) {
                    continue;
                }

                if (!TryApplyCanvasPageImageClip(image, x1, y1, x2, y2)) {
                    ReleasePendingStructureDestination(image.NavigationTarget);
                    images.RemoveAt(i);
                }
            }
        }

        private static bool CanvasClipContainsFrame(PageImage image, double x1, double y1, double x2, double y2) {
            const double tolerance = 0.01D;
            return x1 <= image.X + tolerance
                   && y1 <= image.Y + tolerance
                   && x2 >= image.X + image.W - tolerance
                   && y2 >= image.Y + image.H - tolerance;
        }

        private static bool TryApplyCanvasPageImageClip(PageImage image, double x1, double y1, double x2, double y2) {
            if (image.ClipPath != null && image.ClipPath.Kind == OfficeIMO.Drawing.OfficeClipPathKind.Rectangle) {
                double existingX1 = image.ClipX;
                double existingY1 = image.ClipY + image.ClipHeight - image.ClipPath.Height;
                double existingX2 = image.ClipX + image.ClipPath.Width;
                double existingY2 = image.ClipY + image.ClipHeight;

                x1 = System.Math.Max(x1, existingX1);
                y1 = System.Math.Max(y1, existingY1);
                x2 = System.Math.Min(x2, existingX2);
                y2 = System.Math.Min(y2, existingY2);
            }

            if (x2 <= x1 || y2 <= y1) {
                return false;
            }

            image.ClipPath = OfficeIMO.Drawing.OfficeClipPath.Rectangle(x2 - x1, y2 - y1);
            image.ClipX = x1;
            image.ClipY = y1;
            image.ClipHeight = y2 - y1;
            return true;
        }

        private void ClipCanvasFormFields(System.Collections.Generic.List<FormFieldAnnotation> formFields, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight, OfficeClipPathKind clipKind) {
            if (formFields.Count <= startIndex) return;
            if (clipKind == OfficeClipPathKind.Empty) {
                foreach (var field in formFields.Skip(startIndex)) ReleasePendingStructureDestination(field.NavigationTarget);
                formFields.RemoveRange(startIndex, formFields.Count - startIndex);
                return;
            }
            if (clipKind != OfficeClipPathKind.Rectangle) {
                throw new ArgumentException("Canvas form fields cannot be placed inside nonrectangular clips because PDF widget annotations do not obey page-content clipping.");
            }
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = formFields.Count - 1; i >= startIndex; i--) {
                FormFieldAnnotation formField = formFields[i];
                if (!RectanglesIntersect(formField.X1, formField.Y1, formField.X2, formField.Y2, clipX, clipBottomY, clipRight, clipTop)) {
                    ReleasePendingStructureDestination(formField.NavigationTarget);
                    formFields.RemoveAt(i);
                    continue;
                }
                if (!RectangleContains(clipX, clipBottomY, clipRight, clipTop, formField.X1, formField.Y1, formField.X2, formField.Y2)) {
                    throw new ArgumentException("Canvas form fields must be fully contained by rectangular clips because PDF widget annotations cannot preserve partial clipping.");
                }
                for (int widgetIndex = formField.RadioWidgets.Count - 1; widgetIndex >= 0; widgetIndex--) {
                    RadioButtonWidgetAnnotation widget = formField.RadioWidgets[widgetIndex];
                    if (!RectanglesIntersect(widget.X1, widget.Y1, widget.X2, widget.Y2, clipX, clipBottomY, clipRight, clipTop)) {
                        ReleasePendingStructureDestination(widget.NavigationTarget);
                        formField.RadioWidgets.RemoveAt(widgetIndex);
                        continue;
                    }
                    if (!RectangleContains(clipX, clipBottomY, clipRight, clipTop, widget.X1, widget.Y1, widget.X2, widget.Y2)) {
                        throw new ArgumentException("Canvas radio widgets must be fully contained by rectangular clips because PDF widget annotations cannot preserve partial clipping.");
                    }
                }
            }
        }

        private static bool RectanglesIntersect(double x1, double y1, double x2, double y2, double clipX1, double clipY1, double clipX2, double clipY2) =>
            System.Math.Min(x2, clipX2) > System.Math.Max(x1, clipX1)
            && System.Math.Min(y2, clipY2) > System.Math.Max(y1, clipY1);

        private static bool RectangleContains(double x1, double y1, double x2, double y2, double innerX1, double innerY1, double innerX2, double innerY2) {
            const double tolerance = 0.000001D;
            return innerX1 >= x1 - tolerance
                && innerY1 >= y1 - tolerance
                && innerX2 <= x2 + tolerance
                && innerY2 <= y2 + tolerance;
        }

    }
}
