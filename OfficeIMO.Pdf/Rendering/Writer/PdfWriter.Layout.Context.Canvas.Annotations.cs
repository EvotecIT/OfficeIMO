using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static void TransformCanvasRectangles(System.Collections.Generic.List<LinkAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = annotations.Count - 1; index >= startIndex; index--) {
                LinkAnnotation annotation = annotations[index];
                TransformCanvasRectangle(annotation, transform);
                if (IsEmptyTransformedAnnotation(transform, annotation.X1, annotation.Y1, annotation.X2, annotation.Y2)) annotations.RemoveAt(index);
            }
        }

        private static void TransformCanvasRectangles(System.Collections.Generic.List<TextAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = annotations.Count - 1; index >= startIndex; index--) {
                TextAnnotation annotation = annotations[index];
                TransformCanvasRectangle(annotation, transform);
                if (IsEmptyTransformedAnnotation(transform, annotation.X1, annotation.Y1, annotation.X2, annotation.Y2)) annotations.RemoveAt(index);
            }
        }

        private static void TransformCanvasRectangles(System.Collections.Generic.List<FreeTextAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = annotations.Count - 1; index >= startIndex; index--) {
                FreeTextAnnotation annotation = annotations[index];
                TransformCanvasRectangle(annotation, transform);
                if (IsEmptyTransformedAnnotation(transform, annotation.X1, annotation.Y1, annotation.X2, annotation.Y2)) annotations.RemoveAt(index);
            }
        }

        private static void TransformCanvasRectangles(System.Collections.Generic.List<HighlightAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = annotations.Count - 1; index >= startIndex; index--) {
                HighlightAnnotation annotation = annotations[index];
                TransformCanvasRectangle(annotation, transform);
                if (IsEmptyTransformedAnnotation(transform, annotation.X1, annotation.Y1, annotation.X2, annotation.Y2)) annotations.RemoveAt(index);
            }
        }

        // Singular effects have no painted area even when their bounding line is
        // diagonal. Keep point navigation metadata, but omit invisible annotations.
        private static bool IsEmptyTransformedAnnotation(OfficeTransform transform, double x1, double y1, double x2, double y2) =>
            transform.M11 * transform.M22 == transform.M21 * transform.M12 || x2 <= x1 || y2 <= y1;

        private static void TransformCanvasRectangles(System.Collections.Generic.List<FormFieldAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = startIndex; index < annotations.Count; index++) TransformCanvasRectangle(annotations[index], transform);
        }

        private static void TransformCanvasRectangle(LinkAnnotation annotation, OfficeTransform transform) {
            (annotation.X1, annotation.Y1, annotation.X2, annotation.Y2) = TransformRectangle(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, transform);
        }

        private static void TransformCanvasRectangle(TextAnnotation annotation, OfficeTransform transform) {
            (annotation.X1, annotation.Y1, annotation.X2, annotation.Y2) = TransformRectangle(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, transform);
        }

        private static void TransformCanvasRectangle(FreeTextAnnotation annotation, OfficeTransform transform) {
            (annotation.X1, annotation.Y1, annotation.X2, annotation.Y2) = TransformRectangle(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, transform);
        }

        private static void TransformCanvasRectangle(HighlightAnnotation annotation, OfficeTransform transform) {
            (annotation.X1, annotation.Y1, annotation.X2, annotation.Y2) = TransformRectangle(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, transform);
        }

        private static void TransformCanvasRectangle(FormFieldAnnotation annotation, OfficeTransform transform) {
            if (transform.M11 > 0D && transform.M11 == transform.M22 && transform.M12 == 0D && transform.M21 == 0D) {
                annotation.AppearanceScale *= transform.M11;
            }
            (annotation.X1, annotation.Y1, annotation.X2, annotation.Y2) = TransformRectangle(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, transform);
            for (int index = 0; index < annotation.RadioWidgets.Count; index++) {
                RadioButtonWidgetAnnotation widget = annotation.RadioWidgets[index];
                (widget.X1, widget.Y1, widget.X2, widget.Y2) = TransformRectangle(widget.X1, widget.Y1, widget.X2, widget.Y2, transform);
            }
        }

        private static (double X1, double Y1, double X2, double Y2) TransformRectangle(double x1, double y1, double x2, double y2, OfficeTransform transform) {
            (double left, double top, double right, double bottom) = transform.TransformRectangleBounds(x1, y1, x2 - x1, y2 - y1);
            return (left, top, right, bottom);
        }

        private static void ClipCanvasLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight, OfficeClipPath clipPath) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                LinkAnnotation annotation = annotations[i];
                double x1 = System.Math.Max(annotation.X1, clipX);
                double y1 = System.Math.Max(annotation.Y1, clipBottomY);
                double x2 = System.Math.Min(annotation.X2, clipRight);
                double y2 = System.Math.Min(annotation.Y2, clipTop);
                if (x2 <= x1 || y2 <= y1 || !TryClipCanvasAnnotationRectangle(clipPath, clipX, clipBottomY, clipHeight, ref x1, ref y1, ref x2, ref y2)) {
                    annotations.RemoveAt(i);
                    continue;
                }

                annotation.X1 = x1;
                annotation.Y1 = y1;
                annotation.X2 = x2;
                annotation.Y2 = y2;
            }
        }

        private static void ClipCanvasTextAnnotations(System.Collections.Generic.List<TextAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight, OfficeClipPath clipPath) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                TextAnnotation annotation = annotations[i];
                double x1 = System.Math.Max(annotation.X1, clipX);
                double y1 = System.Math.Max(annotation.Y1, clipBottomY);
                double x2 = System.Math.Min(annotation.X2, clipRight);
                double y2 = System.Math.Min(annotation.Y2, clipTop);
                if (x2 <= x1 || y2 <= y1 || !TryClipCanvasAnnotationRectangle(clipPath, clipX, clipBottomY, clipHeight, ref x1, ref y1, ref x2, ref y2)) {
                    annotations.RemoveAt(i);
                    continue;
                }

                annotation.X1 = x1;
                annotation.Y1 = y1;
                annotation.X2 = x2;
                annotation.Y2 = y2;
            }
        }

        private static void ClipCanvasFreeTextAnnotations(System.Collections.Generic.List<FreeTextAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight, OfficeClipPath clipPath) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                FreeTextAnnotation annotation = annotations[i];
                double x1 = System.Math.Max(annotation.X1, clipX);
                double y1 = System.Math.Max(annotation.Y1, clipBottomY);
                double x2 = System.Math.Min(annotation.X2, clipRight);
                double y2 = System.Math.Min(annotation.Y2, clipTop);
                if (x2 <= x1 || y2 <= y1 || !TryClipCanvasAnnotationRectangle(clipPath, clipX, clipBottomY, clipHeight, ref x1, ref y1, ref x2, ref y2)) {
                    annotations.RemoveAt(i);
                    continue;
                }

                annotation.X1 = x1;
                annotation.Y1 = y1;
                annotation.X2 = x2;
                annotation.Y2 = y2;
            }
        }

        private static void ClipCanvasHighlightAnnotations(System.Collections.Generic.List<HighlightAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight, OfficeClipPath clipPath) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                HighlightAnnotation annotation = annotations[i];
                double x1 = System.Math.Max(annotation.X1, clipX);
                double y1 = System.Math.Max(annotation.Y1, clipBottomY);
                double x2 = System.Math.Min(annotation.X2, clipRight);
                double y2 = System.Math.Min(annotation.Y2, clipTop);
                if (x2 <= x1 || y2 <= y1 || !TryClipCanvasAnnotationRectangle(clipPath, clipX, clipBottomY, clipHeight, ref x1, ref y1, ref x2, ref y2)) {
                    annotations.RemoveAt(i);
                    continue;
                }

                annotation.X1 = x1;
                annotation.Y1 = y1;
                annotation.X2 = x2;
                annotation.Y2 = y2;
            }
        }

    }
}
