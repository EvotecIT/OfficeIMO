namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderTextAnnotationFlowBlock(TextAnnotationBlock annotation) {
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            double spacingBefore = ResolveTopLevelSpacingBefore(annotation.SpacingBefore);
            double needed = spacingBefore + annotation.Height + annotation.SpacingAfter;
            EnsureFixedFlowBlockFits("Text annotation", annotation.Width, needed, contentWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0D) {
                y -= spacingBefore;
            }

            EnsurePage();
            double x = GetAlignedObjectX(currentOpts.MarginLeft, contentWidth, annotation.Width, annotation.Align);
            double bottomY = y - annotation.Height;
            AddTextAnnotation(x, bottomY, annotation.Width, annotation.Height, annotation.Contents, annotation.Icon, annotation.Color, annotation.Open);
            DrawDebugFlowObjectBox(x, bottomY, annotation.Width, annotation.Height);
            y -= annotation.Height + annotation.SpacingAfter;
        }

        private void RenderFreeTextAnnotationFlowBlock(FreeTextAnnotationBlock annotation) {
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            double spacingBefore = ResolveTopLevelSpacingBefore(annotation.SpacingBefore);
            double needed = spacingBefore + annotation.Height + annotation.SpacingAfter;
            EnsureFixedFlowBlockFits("Free text annotation", annotation.Width, needed, contentWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0D) {
                y -= spacingBefore;
            }

            EnsurePage();
            double x = GetAlignedObjectX(currentOpts.MarginLeft, contentWidth, annotation.Width, annotation.Align);
            double bottomY = y - annotation.Height;
            AddFreeTextAnnotation(x, bottomY, annotation.Width, annotation.Height, annotation.Contents, annotation.FontSize, annotation.TextColor, annotation.BorderColor, annotation.BorderWidth, annotation.FillColor, annotation.TextAlign, annotation.Padding, annotation.LineHeight);
            DrawDebugFlowObjectBox(x, bottomY, annotation.Width, annotation.Height);
            y -= annotation.Height + annotation.SpacingAfter;
        }

        private void RenderHighlightAnnotationFlowBlock(HighlightAnnotationBlock annotation) {
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            double spacingBefore = ResolveTopLevelSpacingBefore(annotation.SpacingBefore);
            double needed = spacingBefore + annotation.Height + annotation.SpacingAfter;
            EnsureFixedFlowBlockFits("Highlight annotation", annotation.Width, needed, contentWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0D) {
                y -= spacingBefore;
            }

            EnsurePage();
            double x = GetAlignedObjectX(currentOpts.MarginLeft, contentWidth, annotation.Width, annotation.Align);
            double bottomY = y - annotation.Height;
            AddHighlightAnnotation(x, bottomY, annotation.Width, annotation.Height, annotation.Contents, annotation.Color);
            DrawDebugFlowObjectBox(x, bottomY, annotation.Width, annotation.Height);
            y -= annotation.Height + annotation.SpacingAfter;
        }

        private void RenderCanvasTextAnnotation(PdfCanvasTextAnnotationItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas text annotation");
            double bottomY = currentOpts.PageHeight - item.Y - item.Height;
            AddTextAnnotation(item.X, bottomY, item.Width, item.Height, item.Contents, item.Icon, item.Color, item.Open);
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
        }

        private void RenderCanvasFreeTextAnnotation(PdfCanvasFreeTextAnnotationItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas free text annotation");
            double bottomY = currentOpts.PageHeight - item.Y - item.Height;
            AddFreeTextAnnotation(item.X, bottomY, item.Width, item.Height, item.Contents, item.FontSize, item.TextColor, item.BorderColor, item.BorderWidth, item.FillColor, item.TextAlign, item.Padding, item.LineHeight);
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
        }

        private void RenderCanvasHighlightAnnotation(PdfCanvasHighlightAnnotationItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas highlight annotation");
            double bottomY = currentOpts.PageHeight - item.Y - item.Height;
            AddHighlightAnnotation(item.X, bottomY, item.Width, item.Height, item.Contents, item.Color);
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
        }

        private void AddTextAnnotation(double x, double bottomY, double boxWidth, double boxHeight, string contents, PdfTextAnnotationIcon icon, PdfColor? color, bool open) {
            currentPage!.TextAnnotations.Add(new TextAnnotation {
                X1 = x,
                Y1 = bottomY,
                X2 = x + boxWidth,
                Y2 = bottomY + boxHeight,
                Contents = contents,
                Icon = icon,
                Color = color,
                Open = open
            });
        }

        private void AddFreeTextAnnotation(double x, double bottomY, double boxWidth, double boxHeight, string contents, double fontSize, PdfColor textColor, PdfColor? borderColor, double borderWidth, PdfColor? fillColor, PdfAlign textAlign, double padding, double? lineHeight) {
            currentPage!.FreeTextAnnotations.Add(new FreeTextAnnotation {
                X1 = x,
                Y1 = bottomY,
                X2 = x + boxWidth,
                Y2 = bottomY + boxHeight,
                Contents = contents,
                FontSize = fontSize,
                TextColor = textColor,
                BorderColor = borderColor,
                BorderWidth = borderWidth,
                FillColor = fillColor,
                TextAlign = textAlign,
                Padding = padding,
                LineHeight = lineHeight
            });
        }

        private void AddHighlightAnnotation(double x, double bottomY, double boxWidth, double boxHeight, string contents, PdfColor color) {
            currentPage!.HighlightAnnotations.Add(new HighlightAnnotation {
                X1 = x,
                Y1 = bottomY,
                X2 = x + boxWidth,
                Y2 = bottomY + boxHeight,
                Contents = contents,
                Color = color
            });
        }
    }
}
