using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasBlock(PdfCanvasBlock canvas) {
            EnsurePage();
            foreach (PdfCanvasItem item in canvas.Items) {
                switch (item) {
                    case PdfCanvasStructureItem structure:
                        RenderCanvasStructure(structure);
                        break;
                    case PdfCanvasFigureItem figure:
                        RenderCanvasFigure(figure);
                        break;
                    case PdfCanvasOutlineItem outline:
                        RenderCanvasOutline(outline);
                        break;
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
                    case PdfCanvasTextAnnotationItem textAnnotation:
                        RenderCanvasTextAnnotation(textAnnotation);
                        break;
                    case PdfCanvasFreeTextAnnotationItem freeTextAnnotation:
                        RenderCanvasFreeTextAnnotation(freeTextAnnotation);
                        break;
                    case PdfCanvasHighlightAnnotationItem highlightAnnotation:
                        RenderCanvasHighlightAnnotation(highlightAnnotation);
                        break;
                    case PdfCanvasTableItem table:
                        RenderCanvasTable(table);
                        break;
                    case PdfCanvasClipItem clip:
                        RenderCanvasClip(clip);
                        break;
                    case PdfCanvasEffectItem effect:
                        RenderCanvasEffect(effect);
                        break;
                }
            }
        }

        private void RenderCanvasStructure(PdfCanvasStructureItem item) {
            PdfCanvasStructureOptions options = item.Options;
            int? structureElementIndex = RegisterStructureContainer(
                MapCanvasStructureType(item.Role),
                _canvasStructureParentElementIndex,
                MapCanvasTableHeaderScope(options.HeaderScope),
                options.ColumnSpan,
                options.RowSpan,
                options.AlternativeText);
            int? previous = _canvasStructureParentElementIndex;
            _canvasStructureParentElementIndex = structureElementIndex ?? previous;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _canvasStructureParentElementIndex = previous;
            }
        }

        private static string MapCanvasStructureType(PdfCanvasStructureRole role) {
            if (role == PdfCanvasStructureRole.Section) return "Sect";
            if (role == PdfCanvasStructureRole.Division) return "Div";
            if (role == PdfCanvasStructureRole.Paragraph) return "P";
            if (role == PdfCanvasStructureRole.Heading1) return "H1";
            if (role == PdfCanvasStructureRole.Heading2) return "H2";
            if (role == PdfCanvasStructureRole.Heading3) return "H3";
            if (role == PdfCanvasStructureRole.Heading4) return "H4";
            if (role == PdfCanvasStructureRole.Heading5) return "H5";
            if (role == PdfCanvasStructureRole.Heading6) return "H6";
            if (role == PdfCanvasStructureRole.List) return "L";
            if (role == PdfCanvasStructureRole.ListItem) return "LI";
            if (role == PdfCanvasStructureRole.ListLabel) return "Lbl";
            if (role == PdfCanvasStructureRole.ListBody) return "LBody";
            if (role == PdfCanvasStructureRole.Table) return "Table";
            if (role == PdfCanvasStructureRole.TableRow) return "TR";
            if (role == PdfCanvasStructureRole.TableHeaderCell) return "TH";
            if (role == PdfCanvasStructureRole.TableCell) return "TD";
            return "Caption";
        }

        private static string MapCanvasTableHeaderScope(PdfCanvasTableHeaderScope? scope) {
            if (scope == PdfCanvasTableHeaderScope.Row) return "Row";
            if (scope == PdfCanvasTableHeaderScope.Column) return "Column";
            if (scope == PdfCanvasTableHeaderScope.Both) return "Both";
            return string.Empty;
        }

        private void RenderCanvasFigure(PdfCanvasFigureItem item) {
            EnsurePage();
            int? markedContentId = RegisterFigureStructureElement(item.AlternativeText, _canvasStructureParentElementIndex);
            sb.Append("/Figure << /Alt ")
                .Append(PdfSyntaxEscaper.TextString(item.AlternativeText));
            if (markedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }

            sb.Append(" >> BDC\n");
            bool previous = _suppressCanvasAccessibilityWrappers;
            _suppressCanvasAccessibilityWrappers = true;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _suppressCanvasAccessibilityWrappers = previous;
            }
            sb.Append("EMC\n");
        }

        private void RenderCanvasOutline(PdfCanvasOutlineItem item) {
            EnsurePage();
            currentPage!.Bookmarks.Add(new PageBookmark {
                Level = item.Level,
                Title = item.Title,
                Y = currentOpts.PageHeight - item.Y
            });
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
            string? structureType = _suppressCanvasAccessibilityWrappers ? null : MapCanvasTextStructureType(item.StructureRole);
            int? markedContentId = structureType == null ? null : RegisterTextStructureElement(structureType, _canvasStructureParentElementIndex);
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
                structureType: structureType,
                markedContentId: markedContentId,
                structurePage: currentPage);
            MarkRichFonts(item.Runs);
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
            pageDirty = true;
        }

        private static string MapCanvasTextStructureType(PdfCanvasTextStructureRole role) {
            if (role == PdfCanvasTextStructureRole.Heading1) return "H1";
            if (role == PdfCanvasTextStructureRole.Heading2) return "H2";
            if (role == PdfCanvasTextStructureRole.Heading3) return "H3";
            if (role == PdfCanvasTextStructureRole.Heading4) return "H4";
            if (role == PdfCanvasTextStructureRole.Heading5) return "H5";
            if (role == PdfCanvasTextStructureRole.Heading6) return "H6";
            if (role == PdfCanvasTextStructureRole.Span) return "Span";
            return "P";
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
                if (textContentHeight > textHeight + 0.01D) {
                    item.DiagnosticHandler?.Invoke(new PdfLayoutDiagnostic(
                        PdfLayoutDiagnosticKind.ClippedContent,
                        "PdfCanvasTextBox",
                        "The PDF text box render pass clipped text because wrapped content exceeded the available text area.",
                        item.X,
                        item.Y,
                        item.Width,
                        item.Height));
                }

                double verticalOffset = GetCanvasTextBoxVerticalOffset(style.VerticalAlign, textHeight, textContentHeight);
                var annotations = rotated ? new System.Collections.Generic.List<LinkAnnotation>() : currentPage!.Annotations;
                int? markedContentId = _suppressCanvasAccessibilityWrappers ? null : RegisterTextStructureElement("P", _canvasStructureParentElementIndex);
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
                    structureType: _suppressCanvasAccessibilityWrappers ? null : "P",
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

            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
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
            DrawDebugCanvasItemBox(item.X, topY - block.Shape.Height, block.Shape.Width, block.Shape.Height);
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
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
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
            pageImage.SuppressAccessibilityWrapper = _suppressCanvasAccessibilityWrappers;
            pageImage.RotationAngle = item.RotationAngle;
            pageImage.HorizontalFlip = item.HorizontalFlip;
            pageImage.VerticalFlip = item.VerticalFlip;
            currentPage!.Images.Add(pageImage);
            pageImage.InlineDrawToken = "\n%OIMO_INLINE_IMAGE_" + currentPage.Images.Count.ToString("D6", CultureInfo.InvariantCulture) + "\n";
            sb.Append(pageImage.InlineDrawToken);
            if (!_suppressCanvasAccessibilityWrappers && !string.IsNullOrWhiteSpace(pageImage.AlternativeText)) {
                int? markedContentId = RegisterFigureStructureElement(pageImage.AlternativeText!, _canvasStructureParentElementIndex);
                pageImage.MarkedContentId = markedContentId;
                pageImage.StructElementIndex = FindStructElementIndex(currentPage, markedContentId, "Figure");
            }

            int annotationStart = currentPage!.Annotations.Count;
            AddImageLinkAnnotation(block, imageStyle, pageImage, item.X, bottomY, block.Width, block.Height);
            RotateCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, block.Width, block.Height, item.RotationAngle);
            DrawDebugCanvasItemBox(item.X, bottomY, block.Width, block.Height);
            pageDirty = true;
        }

        private void RenderCanvasClip(PdfCanvasClipItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas clip");
            double bottomY = currentOpts.PageHeight - item.Y - item.Height;
            int annotationStart = currentPage!.Annotations.Count;
            int textAnnotationStart = currentPage.TextAnnotations.Count;
            int freeTextAnnotationStart = currentPage.FreeTextAnnotations.Count;
            int highlightAnnotationStart = currentPage.HighlightAnnotations.Count;
            int imageStart = currentPage.Images.Count;
            int formFieldStart = currentPage.FormFields.Count;
            new ContentStreamBuilder(sb).SaveState();
            AppendClipPath(sb, item.ClipPath, item.X, bottomY, item.Height);

            _canvasClipDepth++;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _canvasClipDepth--;
            }

            ClipCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasTextAnnotations(currentPage.TextAnnotations, textAnnotationStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasFreeTextAnnotations(currentPage.FreeTextAnnotations, freeTextAnnotationStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasHighlightAnnotations(currentPage.HighlightAnnotations, highlightAnnotationStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasPageImages(currentPage.Images, imageStart, item.X, bottomY, item.Width, item.Height);
            ClipCanvasFormFields(currentPage.FormFields, formFieldStart, item.X, bottomY, item.Width, item.Height);
            new ContentStreamBuilder(sb)
                .RestoreState();
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
            pageDirty = true;
        }

        private void RenderCanvasEffect(PdfCanvasEffectItem item) {
            OfficeTransform transform = ConvertTopLeftCanvasTransform(item.Transform, currentOpts.PageHeight);
            int annotationStart = currentPage!.Annotations.Count;
            int textAnnotationStart = currentPage.TextAnnotations.Count;
            int freeTextAnnotationStart = currentPage.FreeTextAnnotations.Count;
            int highlightAnnotationStart = currentPage.HighlightAnnotations.Count;
            int formFieldStart = currentPage.FormFields.Count;
            string? opacityState = EnsureGraphicsState(item.Opacity, item.Opacity);
            int contentStart = sb.Length;
            _canvasClipDepth++;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _canvasClipDepth--;
            }
            string groupContent = sb.ToString(contentStart, sb.Length - contentStart);
            sb.Length = contentStart;
            string token = "\n%OIMO_EFFECT_GROUP_" + (currentPage.EffectGroups.Count + 1).ToString("D6", CultureInfo.InvariantCulture) + "\n";
            currentPage.EffectGroups.Add(new PageEffectGroup {
                Content = groupContent,
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

        private static void TransformCanvasRectangles(System.Collections.Generic.List<LinkAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = startIndex; index < annotations.Count; index++) TransformCanvasRectangle(annotations[index], transform);
        }

        private static void TransformCanvasRectangles(System.Collections.Generic.List<TextAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = startIndex; index < annotations.Count; index++) TransformCanvasRectangle(annotations[index], transform);
        }

        private static void TransformCanvasRectangles(System.Collections.Generic.List<FreeTextAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = startIndex; index < annotations.Count; index++) TransformCanvasRectangle(annotations[index], transform);
        }

        private static void TransformCanvasRectangles(System.Collections.Generic.List<HighlightAnnotation> annotations, int startIndex, OfficeTransform transform) {
            for (int index = startIndex; index < annotations.Count; index++) TransformCanvasRectangle(annotations[index], transform);
        }

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
            (annotation.X1, annotation.Y1, annotation.X2, annotation.Y2) = TransformRectangle(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, transform);
        }

        private static (double X1, double Y1, double X2, double Y2) TransformRectangle(double x1, double y1, double x2, double y2, OfficeTransform transform) {
            (double left, double top, double right, double bottom) = transform.TransformRectangleBounds(x1, y1, x2 - x1, y2 - y1);
            return (left, top, right, bottom);
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

        private static void ClipCanvasTextAnnotations(System.Collections.Generic.List<TextAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                TextAnnotation annotation = annotations[i];
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

        private static void ClipCanvasFreeTextAnnotations(System.Collections.Generic.List<FreeTextAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                FreeTextAnnotation annotation = annotations[i];
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

        private static void ClipCanvasHighlightAnnotations(System.Collections.Generic.List<HighlightAnnotation> annotations, int startIndex, double clipX, double clipBottomY, double clipWidth, double clipHeight) {
            double clipRight = clipX + clipWidth;
            double clipTop = clipBottomY + clipHeight;
            for (int i = annotations.Count - 1; i >= startIndex; i--) {
                HighlightAnnotation annotation = annotations[i];
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

                if (!string.IsNullOrEmpty(image.InlineDrawToken) || CanvasClipContainsFrame(image, x1, y1, x2, y2)) {
                    continue;
                }

                if (!TryApplyCanvasPageImageClip(image, x1, y1, x2, y2)) {
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
