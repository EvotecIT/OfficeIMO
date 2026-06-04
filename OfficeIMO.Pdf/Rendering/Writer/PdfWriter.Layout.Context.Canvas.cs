using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasBlock(PdfCanvasBlock canvas) {
            EnsurePage();
            foreach (PdfCanvasItem item in canvas.Items) {
                switch (item) {
                    case PdfCanvasTextItem text:
                        RenderCanvasText(text);
                        break;
                    case PdfCanvasTextBoxItem textBox:
                        RenderCanvasTextBox(textBox);
                        break;
                    case PdfCanvasShapeItem shape:
                        RenderCanvasShape(shape);
                        break;
                    case PdfCanvasDrawingItem drawing:
                        RenderCanvasDrawing(drawing);
                        break;
                    case PdfCanvasImageItem image:
                        RenderCanvasImage(image);
                        break;
                    case PdfCanvasTableItem table:
                        RenderCanvasTable(table);
                        break;
                    case PdfCanvasClipItem clip:
                        RenderCanvasClip(clip);
                        break;
                }
            }
        }

        private void RenderCanvasText(PdfCanvasTextItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas text");
            double size = item.FontSize ?? currentOpts.DefaultFontSize;
            double leading = item.LineHeight ?? size * 1.2D;
            var block = new RichParagraphBlock(item.Runs, item.Align, item.DefaultColor);
            var wrap = WrapRichRunsCore(item.Runs, item.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
            if (wrap.Lines.Count == 0) {
                return;
            }

            double topY = currentOpts.PageHeight - item.Y;
            double bottomY = topY - item.Height;
            int? markedContentId = RegisterTextStructureElement("P");
            WriteClippedRichParagraph(
                sb,
                block,
                wrap.Lines,
                wrap.LineHeights,
                currentOpts,
                FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), size, topY),
                size,
                leading,
                currentPage!.Annotations,
                item.X,
                bottomY,
                item.Width,
                item.Height,
                item.X,
                item.Width,
                structureType: "P",
                markedContentId: markedContentId,
                structurePage: currentPage);
            MarkRichFonts(item.Runs);
            pageDirty = true;
        }

        private void RenderCanvasTextBox(PdfCanvasTextBoxItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas text box");
            PdfCanvasTextBoxStyle style = item.Style;
            double topY = currentOpts.PageHeight - item.Y;
            double bottomY = topY - item.Height;
            bool rotated = item.RotationAngle != 0D;
            if (rotated) {
                BeginRotatedCanvasFrame(item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            }

            if (style.Background.HasValue || (style.BorderColor.HasValue && style.BorderWidth > 0D)) {
                OfficeIMO.Drawing.OfficeShape shape = style.CornerRadius > 0D
                    ? OfficeIMO.Drawing.OfficeShape.RoundedRectangle(item.Width, item.Height, style.CornerRadius)
                    : OfficeIMO.Drawing.OfficeShape.Rectangle(item.Width, item.Height);
                shape.FillColor = style.Background?.ToOfficeColor();
                shape.FillOpacity = style.Background.HasValue ? style.BackgroundOpacity : null;
                shape.StrokeColor = style.BorderWidth > 0D ? style.BorderColor?.ToOfficeColor() : null;
                shape.StrokeWidth = style.BorderWidth;
                shape.StrokeDashStyle = style.BorderDashStyle;
                shape.StrokeLineCap = style.BorderLineCap;
                shape.StrokeLineJoin = style.BorderLineJoin;

                ShapeBlock block = PdfDocument.CreateShapeBlock(shape, PdfAlign.Left, spacingBefore: 0D, spacingAfter: 0D);
                PdfDrawingStyle drawingStyle = ResolveDrawingStyle(block, currentOpts);
                PdfDocument.ValidateDrawingStyle(drawingStyle, "Canvas text box");
                DrawShapeAt(block, drawingStyle, item.X, item.Width, topY);
            }

            double paddingLeft = style.EffectivePaddingLeft;
            double paddingRight = style.EffectivePaddingRight;
            double paddingTop = style.EffectivePaddingTop;
            double paddingBottom = style.EffectivePaddingBottom;
            double textX = item.X + paddingLeft;
            double textWidth = item.Width - paddingLeft - paddingRight;
            double textHeight = item.Height - paddingTop - paddingBottom;
            double textTopY = topY - paddingTop;
            double textBottomY = bottomY + paddingBottom;
            double size = style.FontSize ?? currentOpts.DefaultFontSize;
            double leading = style.LineHeight ?? size * 1.2D;
            PdfStandardFont baseFont = ChooseNormal(style.Font ?? currentOpts.DefaultFont);
            var blockText = new RichParagraphBlock(item.Runs, style.Align, style.TextColor);
            var wrap = WrapRichRunsCore(item.Runs, textWidth, size, baseFont, leading, null, DefaultParagraphTabStopWidth, currentOpts);
            if (wrap.Lines.Count > 0) {
                double textContentHeight = MeasureRichLinesHeight(wrap.LineHeights, wrap.Lines.Count, leading);
                double verticalOffset = GetCanvasTextBoxVerticalOffset(style.VerticalAlign, textHeight, textContentHeight);
                var annotations = rotated ? new System.Collections.Generic.List<LinkAnnotation>() : currentPage!.Annotations;
                int? markedContentId = RegisterTextStructureElement("P");
                WriteClippedRichParagraph(
                    sb,
                    blockText,
                    wrap.Lines,
                    wrap.LineHeights,
                    currentOpts,
                    FirstTextBaselineFromTop(baseFont, size, textTopY - verticalOffset),
                    size,
                    leading,
                    annotations,
                    textX,
                    textBottomY,
                    textWidth,
                    textHeight,
                    textX,
                    textWidth,
                    structureType: "P",
                    markedContentId: markedContentId,
                    structurePage: currentPage);
                MarkRichFonts(item.Runs);
                if (rotated && annotations.Count > 0) {
                    RotateCanvasLinkAnnotations(annotations, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
                    currentPage!.Annotations.AddRange(annotations);
                }
            }

            if (rotated) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }

            pageDirty = true;
        }

        private static double GetCanvasTextBoxVerticalOffset(PdfVerticalAlign align, double boxHeight, double contentHeight) {
            double unusedHeight = System.Math.Max(0D, boxHeight - contentHeight);
            return align switch {
                PdfVerticalAlign.Middle => unusedHeight / 2D,
                PdfVerticalAlign.Bottom => unusedHeight,
                _ => 0D
            };
        }

        private void RenderCanvasShape(PdfCanvasShapeItem item) {
            ShapeBlock block = item.Block;
            ValidateCanvasBox(item.X, item.Y, block.Shape.Width, block.Shape.Height, "Canvas shape");
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Canvas shape");
            double topY = currentOpts.PageHeight - item.Y;
            int annotationStart = currentPage!.Annotations.Count;
            int? structElementIndex = DrawShapeAt(block, style, item.X, block.Shape.Width, topY);
            if (!string.IsNullOrEmpty(block.LinkUri)) {
                currentPage.Annotations.Add(new LinkAnnotation {
                    X1 = item.X,
                    Y1 = topY - block.Shape.Height,
                    X2 = item.X + block.Shape.Width,
                    Y2 = topY,
                    Uri = block.LinkUri!,
                    Contents = block.LinkContents,
                    StructElementIndex = structElementIndex
                });
            }

            RotateCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, topY - block.Shape.Height, block.Shape.Width, block.Shape.Height, item.RotationAngle);
            pageDirty = true;
        }

        private void RenderCanvasDrawing(PdfCanvasDrawingItem item) {
            DrawingBlock block = item.Block;
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas drawing");
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Canvas drawing");
            double topY = currentOpts.PageHeight - item.Y;
            double bottomY = topY - item.Height;
            int annotationStart = currentPage!.Annotations.Count;
            bool rotated = item.RotationAngle != 0D;
            if (rotated) {
                BeginRotatedCanvasFrame(item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            }

            bool markedContent;
            int? structElementIndex = AppendDrawingMarkedContentBegin(style, out markedContent);
            double scaleX = item.Width / block.Drawing.Width;
            double scaleY = item.Height / block.Drawing.Height;
            bool scaled = Math.Abs(scaleX - 1D) > 0.0001D || Math.Abs(scaleY - 1D) > 0.0001D;
            if (scaled) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .TransformMatrix(scaleX, 0D, 0D, scaleY, item.X, bottomY);
            }

            for (int i = 0; i < block.Drawing.Elements.Count; i++) {
                if (block.Drawing.Elements[i] is OfficeIMO.Drawing.OfficeDrawingShape drawingShape) {
                    double xShape = scaled ? drawingShape.X : item.X + drawingShape.X;
                    double shapeBottomY = scaled
                        ? block.Drawing.Height - drawingShape.Y - drawingShape.Shape.Height
                        : topY - drawingShape.Y - drawingShape.Shape.Height;
                    DrawShapeGeometryAt(drawingShape.Shape, xShape, shapeBottomY);
                } else if (block.Drawing.Elements[i] is OfficeIMO.Drawing.OfficeDrawingText drawingText) {
                    double textX = scaled ? drawingText.X : item.X + drawingText.X;
                    double textTopY = scaled ? block.Drawing.Height - drawingText.Y : topY - drawingText.Y;
                    DrawDrawingTextAt(drawingText, textX, textTopY);
                }
            }

            if (scaled) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }

            AppendDrawingMarkedContentEnd(markedContent);
            if (!string.IsNullOrEmpty(block.LinkUri)) {
                currentPage.Annotations.Add(new LinkAnnotation {
                    X1 = item.X,
                    Y1 = bottomY,
                    X2 = item.X + item.Width,
                    Y2 = topY,
                    Uri = block.LinkUri!,
                    Contents = block.LinkContents,
                    StructElementIndex = structElementIndex
                });
            }

            if (rotated) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }

            RotateCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            pageDirty = true;
        }

        private void RenderCanvasImage(PdfCanvasImageItem item) {
            ImageBlock block = item.Block;
            ValidateCanvasBox(item.X, item.Y, block.Width, block.Height, "Canvas image");
            PdfImageStyle imageStyle = ResolveImageStyle(block, currentOpts);
            PdfDocument.ValidateImageStyleForBox(imageStyle, block.Width, block.Height, nameof(imageStyle.ClipPath));
            PdfDocument.ValidateImageFitDimensions(block.Info, imageStyle.Fit, nameof(imageStyle.Fit));
            double bottomY = currentOpts.PageHeight - item.Y - block.Height;
            PageImage pageImage = CreatePageImage(block, imageStyle, item.X, bottomY, block.Width, block.Height);
            pageImage.RotationAngle = item.RotationAngle;
            pageImage.HorizontalFlip = item.HorizontalFlip;
            pageImage.VerticalFlip = item.VerticalFlip;
            currentPage!.Images.Add(pageImage);
            pageImage.InlineDrawToken = "\n%OIMO_INLINE_IMAGE_" + currentPage.Images.Count.ToString("D6", CultureInfo.InvariantCulture) + "\n";
            sb.Append(pageImage.InlineDrawToken);
            if (!string.IsNullOrWhiteSpace(pageImage.AlternativeText)) {
                int? markedContentId = RegisterFigureStructureElement(pageImage.AlternativeText!);
                pageImage.MarkedContentId = markedContentId;
                pageImage.StructElementIndex = FindStructElementIndex(currentPage, markedContentId, "Figure");
            }

            int annotationStart = currentPage!.Annotations.Count;
            AddImageLinkAnnotation(block, imageStyle, pageImage, item.X, bottomY, block.Width, block.Height);
            RotateCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, block.Width, block.Height, item.RotationAngle);
            pageDirty = true;
        }

        private void RenderCanvasClip(PdfCanvasClipItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas clip");
            double bottomY = currentOpts.PageHeight - item.Y - item.Height;
            int annotationStart = currentPage!.Annotations.Count;
            int imageStart = currentPage.Images.Count;
            int formFieldStart = currentPage.FormFields.Count;
            new ContentStreamBuilder(sb)
                .SaveState()
                .Rectangle(item.X, bottomY, item.Width, item.Height)
                .ClipPath()
                .EndPath();

            _canvasClipDepth++;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _canvasClipDepth--;
            }

            ClipCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasPageImages(currentPage.Images, imageStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasFormFields(currentPage.FormFields, formFieldStart, item.X, bottomY, item.Width, item.Height);
            new ContentStreamBuilder(sb)
                .RestoreState();
            pageDirty = true;
        }

        private void ValidateCanvasBox(double x, double yFromTop, double boxWidth, double boxHeight, string name) {
            if (double.IsNaN(x) || double.IsNaN(yFromTop) || double.IsInfinity(x) || double.IsInfinity(yFromTop)) {
                throw new ArgumentOutOfRangeException(name, name + " coordinates must be finite.");
            }

            if (double.IsNaN(boxWidth) || double.IsNaN(boxHeight) || double.IsInfinity(boxWidth) || double.IsInfinity(boxHeight) || boxWidth < 0D || boxHeight < 0D || (boxWidth == 0D && boxHeight == 0D)) {
                throw new ArgumentOutOfRangeException(name, name + " dimensions must be finite non-negative values with at least one positive dimension.");
            }

            if (_canvasClipDepth > 0) {
                return;
            }

            if (x + boxWidth > currentOpts.PageWidth + 0.001D || yFromTop + boxHeight > currentOpts.PageHeight + 0.001D) {
                throw new ArgumentException(name + " exceeds the current page bounds.");
            }
        }

        private static void ClipCanvasLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                LinkAnnotation annotation = annotations[i];
                double x1 = System.Math.Max(annotation.X1, clipX);
                double y1 = System.Math.Max(annotation.Y1, clipBottomY);
                double x2 = System.Math.Min(annotation.X2, clipRight);
                double y2 = System.Math.Min(annotation.Y2, clipTop);
                if (x2 <= x1 || y2 <= y1) {
                    annotations.RemoveAt(i);
                    continue;
                }

                annotation.X1 = x1;
                annotation.Y1 = y1;
                annotation.X2 = x2;
                annotation.Y2 = y2;
            }
        }

        private void ClipCanvasPageImages(System.Collections.Generic.List<PageImage> images, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = images.Count - 1; i >= startIndex; i--) {
                PageImage image = images[i];
                double x1 = System.Math.Max(image.X, clipX);
                double y1 = System.Math.Max(image.Y, clipBottomY);
                double x2 = System.Math.Min(image.X + image.W, clipRight);
                double y2 = System.Math.Min(image.Y + image.H, clipTop);
                if (x2 <= x1 || y2 <= y1) {
                    if (!string.IsNullOrEmpty(image.InlineDrawToken)) {
                        sb.Replace(image.InlineDrawToken, string.Empty);
                    }

                    images.RemoveAt(i);
                    continue;
                }

                image.ClipPath = OfficeIMO.Drawing.OfficeClipPath.Rectangle(x2 - x1, y2 - y1);
                image.ClipX = x1;
                image.ClipY = y1;
                image.ClipHeight = y2 - y1;
            }
        }

        private static void ClipCanvasFormFields(System.Collections.Generic.List<FormFieldAnnotation> formFields, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = formFields.Count - 1; i >= startIndex; i--) {
                FormFieldAnnotation formField = formFields[i];
                double x1 = System.Math.Max(formField.X1, clipX);
                double y1 = System.Math.Max(formField.Y1, clipBottomY);
                double x2 = System.Math.Min(formField.X2, clipRight);
                double y2 = System.Math.Min(formField.Y2, clipTop);
                if (x2 <= x1 || y2 <= y1) {
                    formFields.RemoveAt(i);
                    continue;
                }

                formField.X1 = x1;
                formField.Y1 = y1;
                formField.X2 = x2;
                formField.Y2 = y2;
            }
        }

        private void BeginRotatedCanvasFrame(double x, double bottomY, double width, double height, double rotationAngle) {
            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            double e = centerX - cos * centerX + sin * centerY;
            double f = centerY - sin * centerX - cos * centerY;

            new ContentStreamBuilder(sb)
                .SaveState()
                .TransformMatrix(cos, sin, -sin, cos, e, f);
        }

        private static void RotateCanvasLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annotations, double x, double bottomY, double width, double height, double rotationAngle) {
            RotateCanvasLinkAnnotations(annotations, 0, x, bottomY, width, height, rotationAngle);
        }

        private static void RotateCanvasLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annotations, int startIndex, double x, double bottomY, double width, double height, double rotationAngle) {
            if (rotationAngle == 0D) {
                return;
            }

            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            for (int i = startIndex; i < annotations.Count; i++) {
                RotateCanvasLinkAnnotation(annotations[i], centerX, centerY, cos, sin);
            }
        }

        private static void RotateCanvasPageImages(System.Collections.Generic.List<PageImage> images, int startIndex, double x, double bottomY, double width, double height, double rotationAngle) {
            if (rotationAngle == 0D) {
                return;
            }

            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            for (int i = startIndex; i < images.Count; i++) {
                PageImage image = images[i];
                double imageCenterX = image.X + image.W / 2D;
                double imageCenterY = image.Y + image.H / 2D;
                RotateCanvasPoint(imageCenterX, imageCenterY, centerX, centerY, cos, sin, out double rotatedCenterX, out double rotatedCenterY);
                image.X = rotatedCenterX - image.W / 2D;
                image.Y = rotatedCenterY - image.H / 2D;
                image.RotationAngle += rotationAngle;
            }
        }

        private static void RotateCanvasFormFields(System.Collections.Generic.List<FormFieldAnnotation> formFields, int startIndex, double x, double bottomY, double width, double height, double rotationAngle) {
            if (rotationAngle == 0D) {
                return;
            }

            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            for (int i = startIndex; i < formFields.Count; i++) {
                RotateCanvasFormField(formFields[i], centerX, centerY, cos, sin);
            }
        }

        private static void RotateCanvasLinkAnnotation(LinkAnnotation annotation, double centerX, double centerY, double cos, double sin) {
            RotateCanvasPoint(annotation.X1, annotation.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
            RotateCanvasPoint(annotation.X1, annotation.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
            RotateCanvasPoint(annotation.X2, annotation.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
            RotateCanvasPoint(annotation.X2, annotation.Y2, centerX, centerY, cos, sin, out double x4, out double y4);

            annotation.X1 = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            annotation.Y1 = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            annotation.X2 = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            annotation.Y2 = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
        }

        private static void RotateCanvasFormField(FormFieldAnnotation formField, double centerX, double centerY, double cos, double sin) {
            RotateCanvasPoint(formField.X1, formField.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
            RotateCanvasPoint(formField.X1, formField.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
            RotateCanvasPoint(formField.X2, formField.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
            RotateCanvasPoint(formField.X2, formField.Y2, centerX, centerY, cos, sin, out double x4, out double y4);

            formField.X1 = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            formField.Y1 = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            formField.X2 = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            formField.Y2 = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
        }

        private static void RotateCanvasPoint(double x, double y, double centerX, double centerY, double cos, double sin, out double rotatedX, out double rotatedY) {
            double dx = x - centerX;
            double dy = y - centerY;
            rotatedX = centerX + cos * dx - sin * dy;
            rotatedY = centerY + sin * dx + cos * dy;
        }
    }
}
