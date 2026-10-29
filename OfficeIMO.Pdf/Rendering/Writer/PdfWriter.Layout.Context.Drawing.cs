using System;
using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static global::OfficeIMO.OfficeColor? ToPdfColor(OfficeIMO.OfficeColor? color) =>
            color.HasValue ? (color.Value.A == 0 ? (global::OfficeIMO.OfficeColor?)null : color.Value) : null;

        private string? EnsureGraphicsState(double fillOpacity, double strokeOpacity, OfficeBlendMode blendMode = OfficeBlendMode.Normal) {
            if (fillOpacity >= 1D && strokeOpacity >= 1D && blendMode == OfficeBlendMode.Normal) {
                return null;
            }

            EnsurePage();
            for (int i = 0; i < currentPage!.GraphicsStates.Count; i++) {
                var existing = currentPage.GraphicsStates[i];
                if (existing.FillOpacity.Equals(fillOpacity) &&
                    existing.StrokeOpacity.Equals(strokeOpacity) &&
                    existing.BlendMode == blendMode) {
                    return existing.Name;
                }
            }

            string name = "GS" + (currentPage.GraphicsStates.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.GraphicsStates.Add(new PageGraphicsState {
                Name = name,
                FillOpacity = fillOpacity,
                StrokeOpacity = strokeOpacity,
                BlendMode = blendMode
            });
            return name;
        }

        private string? EnsureOpacityState(OfficeIMO.Drawing.OfficeShape shape) {
            var opacity = ResolveShapePaintOpacity(shape);
            return EnsureGraphicsState(opacity.Fill, opacity.Stroke);
        }

        private string? EnsureFillGradient(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
            if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) return null;
            EnsurePage();
            return EnsureShapeFillGradient(currentPage!.Shadings, shape, xShape, bottomY, localCoordinates, _drawingIsOpacityData);
        }

        private void DrawShapeShadowAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            var shadow = shape.Shadow;
            if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) return;
            double coreOpacity = shadow.Opacity * shadow.Color.A / 255D;
            global::OfficeIMO.OfficeColor shadowColor = global::OfficeIMO.OfficeColor.FromRgb(shadow.Color.R, shadow.Color.G, shadow.Color.B);
            double shadowX = xShape + shadow.OffsetX;
            double shadowBottomY = bottomY - shadow.OffsetY;
            ResolveShadowGeometry(shape, out bool hasFill, out bool hasStroke);
            IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
                coreOpacity,
                shadow.BlurRadius,
                shape.StrokeWidth,
                hasFill,
                hasStroke,
                OfficeShadowLayerPlanner.CanExpand(shape),
                Math.Min(shape.Width, shape.Height));
            for (int index = 0; index < layers.Count; index++) {
                OfficeShadowLayer layer = layers[index];
                OfficeShape layerShape = Math.Abs(layer.Expansion) > 0.000000001D
                    ? OfficeShadowLayerPlanner.CreateExpandedShape(shape, layer.Expansion)
                    : shape;
                DrawShapeShadowLayerAt(
                    layerShape,
                    shadowColor,
                    shadowX - layer.Expansion,
                    shadowBottomY - layer.Expansion,
                    layer.StrokeWidth,
                    layer.Opacity,
                    layer.HasFill,
                    layer.HasStroke);
            }
            pageDirty = true;
        }

        private void DrawShapeShadowLayerAt(
            OfficeShape shape,
            global::OfficeIMO.OfficeColor color,
            double x,
            double bottomY,
            double strokeWidth,
            double opacity,
            bool hasFill,
            bool hasStroke) {
            var content = new ContentStreamBuilder(sb).SaveState();
            string? graphicsState = EnsureGraphicsState(opacity, opacity);
            if (graphicsState != null) content.GraphicsState(graphicsState);
            DrawShapeShadowLayer(sb, shape, color, x, bottomY, strokeWidth, hasFill, hasStroke);
            content.RestoreState();
        }

        private void DrawShapeGeometryAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            if (shape.ClipPath?.Kind == OfficeClipPathKind.Empty) return;

            DrawShapeGlowAt(shape, xShape, bottomY);
            DrawShapeShadowAt(shape, xShape, bottomY);

            if (shape.StrokeWidth > 0D && (shape.StrokeGradient != null || shape.StrokeRadialGradient != null)) {
                OfficeShape fill = shape.Clone();
                fill.StrokeColor = null; fill.StrokeGradient = null; fill.StrokeRadialGradient = null; fill.StrokeWidth = 0;
                fill.Shadow = null;
                fill.Glow = null;
                DrawShapeGeometryAt(fill, xShape, bottomY);
                OfficeShape? stroke = CreateGradientStrokeShape(shape);
                if (stroke != null) DrawShapeGeometryAt(stroke, xShape, bottomY);
                DrawShapeMarkersAt(shape, xShape, bottomY);
                return;
            }

            string? opacityState = EnsureOpacityState(shape);
            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .GraphicsState(opacityState);
            }

            if (shape.Transform.HasValue || RequiresExactStrokeWriter(shape)) {
                pageDirty = true;
                OfficeIMO.Drawing.OfficeShape renderShape = shape;
                if (!renderShape.Transform.HasValue) {
                    renderShape = shape.Clone();
                    renderShape.Transform = OfficeTransform.Identity;
                }
                string? shadingName = EnsureFillGradient(renderShape, xShape, bottomY, localCoordinates: true);
                DrawTransformedShape(sb, renderShape, shadingName == null ? ToPdfColor(renderShape.FillColor) : null, ToPdfColor(renderShape.StrokeColor), shadingName, xShape, bottomY);
            } else {
                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .SaveState();
                    AppendClipPath(sb, shape.ClipPath, xShape, bottomY, shape.Height);
                }

                string? shadingName = EnsureFillGradient(shape, xShape, bottomY, localCoordinates: false);
                if (shadingName != null) {
                    pageDirty = true;
                    DrawGradientShape(sb, shape, shadingName, xShape, bottomY);
                }

                global::OfficeIMO.OfficeColor? fillColor = shadingName == null ? ToPdfColor(shape.FillColor) : null;
                if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                    pageDirty = true;
                    DrawLine(sb, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle) {
                    pageDirty = true;
                    DrawRoundedRectangle(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height, shape.CornerRadius);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle) {
                    pageDirty = true;
                    DrawRectangle(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Ellipse) {
                    pageDirty = true;
                    DrawEllipse(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Polygon) {
                    pageDirty = true;
                    DrawPolygon(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height, shape.FillRule);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                    pageDirty = true;
                    DrawPath(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, xShape, bottomY, shape.Height, shape.FillRule);
                }

                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .RestoreState();
                }
            }

            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }
            DrawShapeMarkersAt(shape, xShape, bottomY);
        }

        private int? DrawShapeAt(ShapeBlock block, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            double xShape = GetAlignedObjectX(containerX, containerWidth, block.Shape.Width, style.Align);
            bool markedContent;
            int? structElementIndex = AppendDrawingMarkedContentBegin(style, out markedContent);
            DrawShapeGeometryAt(block.Shape, xShape, topY - block.Shape.Height);
            AppendDrawingMarkedContentEnd(markedContent);
            return structElementIndex;
        }

        private int? DrawDrawingAt(DrawingBlock block, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            OfficeDrawing drawing = ResolveDrawingFlowScene(block, style, containerWidth);
            double xDrawing = GetAlignedObjectX(containerX, containerWidth, drawing.Width, style.Align);
            bool markedContent;
            int? structElementIndex = AppendDrawingMarkedContentBegin(
                style,
                out markedContent,
                recordDrawingEvidence: !ChildImagesOwnDrawingAccessibility(block.Drawing, style));
            bool previousSuppressAccessibilityWrappers = _suppressCanvasAccessibilityWrappers;
            if (markedContent || style.Decorative) {
                _suppressCanvasAccessibilityWrappers = true;
            }
            try {
                DrawDrawingElements(drawing, xDrawing, topY);
            } finally {
                _suppressCanvasAccessibilityWrappers = previousSuppressAccessibilityWrappers;
                AppendDrawingMarkedContentEnd(markedContent);
            }
            return structElementIndex;
        }

        private void DrawDrawingElements(OfficeDrawing drawing, double originX, double originTopY, OfficeDrawingTextMetrics? textMetrics = null) {
            textMetrics ??= CreateDrawingTextMetrics(currentOpts, cancellationToken);
            for (int i = 0; i < drawing.Elements.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeDrawingElement element = drawing.Elements[i];
                DrawSourceStructuredElement(element, () => DrawDrawingElement(element, originX, originTopY, textMetrics));
            }
        }

        private void DrawDrawingElement(OfficeDrawingElement element, double originX, double originTopY, OfficeDrawingTextMetrics textMetrics) {
                if (element is OfficeDrawingShape shape) {
                    double xShape = originX + shape.X;
                    double bottomY = originTopY - shape.Y - shape.Shape.Height;
                    DrawShapeGeometryAt(shape.Shape, xShape, bottomY);
                } else if (element is OfficeDrawingLink link) {
                    DrawDrawingLinkAt(link, originX, originTopY);
                } else if (element is OfficeDrawingText text) {
                    DrawDrawingTextAt(text, originX, originTopY, textMetrics);
                } else if (element is OfficeDrawingRichText richText) {
                    DrawDrawingRichTextAt(richText, originX, originTopY, textMetrics);
                } else if (element is OfficeDrawingImage image) {
                    DrawDrawingImageAt(image, originX, originTopY);
                } else if (element is OfficeDrawingGroup group) {
                    DrawDrawingGroupAt(group, originX, originTopY, textMetrics);
                } else if (element is OfficeDrawingEffectGroup effectGroup) {
                    DrawDrawingEffectAt(effectGroup, originX, originTopY, textMetrics);
                } else if (element is OfficeDrawingTilingPattern pattern) {
                    DrawDrawingPatternAt(pattern, originX, originTopY, textMetrics);
                } else {
                    throw new NotSupportedException(
                        "OfficeIMO.Pdf does not yet support drawing elements of type " +
                        element.GetType().Name + ".");
                }
        }

        private void DrawDrawingGroupAt(OfficeDrawingGroup group, double originX, double originTopY, OfficeDrawingTextMetrics textMetrics) {
            void DrawGroupContent() {
                int linkStart = currentPage!.Annotations.Count;
                double clipX = originX + group.X;
                double clipBottomY = originTopY - group.Y - group.ClipPath.Height;
                new ContentStreamBuilder(sb).SaveState();
                AppendClipPath(sb, group.ClipPath, clipX, clipBottomY, group.ClipPath.Height);
                DrawDrawingElements(
                    group.InnerDrawing,
                    clipX + group.ContentOffsetX,
                    originTopY - group.Y - group.ContentOffsetY, textMetrics);
                ClipCanvasLinkAnnotations(currentPage.Annotations, linkStart, clipX, clipBottomY, group.ClipPath.Width, group.ClipPath.Height, group.ClipPath);
                new ContentStreamBuilder(sb).RestoreState();
            }

            void DrawGroupPaint() {
                if (group.FrameTransform.HasValue && group.FrameTransform.Value.HasTransform) {
                    OfficeTransform pageTransform = ToTopLeftPageTransform(
                        group.FrameTransform.Value.CreateDestinationTransform(),
                        originX,
                        originTopY);
                    RenderEffectGroup(pageTransform, 1D, DrawGroupContent);
                } else {
                    DrawGroupContent();
                }
            }

            if (group.ActualText == null || _suppressCanvasActualTextChildren || _suppressCanvasAccessibilityWrappers) {
                DrawGroupPaint();
            } else {
                RenderLogicalText(
                    group.ActualText,
                    originX + group.ActualTextAnchorX,
                    originTopY - group.ActualTextAnchorY,
                    DrawGroupPaint);
            }
        }

        private OfficeTransform ToTopLeftPageTransform(OfficeTransform localTransform, double originX, double originTopY) {
            double pageTopY = currentOpts.PageHeight - originTopY;
            return OfficeTransform
                .Translate(-originX, -pageTopY)
                .Then(localTransform)
                .Then(OfficeTransform.Translate(originX, pageTopY));
        }

        private void DrawDrawingImageAt(OfficeDrawingImage image, double originX, double originTopY, long? logicalOrder = null) {
            OfficeImageProjection projection = image.Projection;
            if (TryDrawSvgImageAt(image, originX, originTopY, logicalOrder)) return;
            PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(image.EncodedBytes);
            var imageStyle = new PdfImageStyle {
                Fit = OfficeImageFit.Stretch,
                RotationAngle = -projection.RotationDegrees,
                AlternativeText = image.AlternativeText
            };
            if (projection.HasCrop) {
                imageStyle.SourceCrop = new PdfImageSourceCrop(
                    projection.SourceCrop.Left,
                    projection.SourceCrop.Top,
                    projection.SourceCrop.Right,
                    projection.SourceCrop.Bottom);
            }

            double targetX = originX + projection.X;
            double targetBottomY = originTopY - projection.Y - projection.Height;
            var block = new ImageBlock(
                prepared.Data,
                projection.Width,
                projection.Height,
                prepared.Info,
                imageStyle,
                useDataSnapshot: true,
                preparedStream: prepared.PreparedStream);
            PageImage pageImage = CreatePageImage(
                block,
                imageStyle,
                targetX,
                targetBottomY,
                projection.Width,
                projection.Height);
            pageImage.Interpolate = image.Interpolate;
            pageImage.IsOpacityData = _drawingIsOpacityData;
            pageImage.Opacity = image.Opacity;
            pageImage.GraphicsStateName = EnsureGraphicsState(image.Opacity, image.Opacity);
            pageImage.HorizontalFlip = projection.FlipHorizontal;
            pageImage.VerticalFlip = projection.FlipVertical;
            pageImage.RotationCenterX = originX + projection.RotationCenterX;
            pageImage.RotationCenterY = originTopY - projection.RotationCenterY;
            pageImage.SuppressAccessibilityWrapper = _suppressCanvasAccessibilityWrappers;
            pageImage.StructureParentElement = _canvasStructureParentElement;
            pageImage.LogicalOrder = logicalOrder;
            ReserveImageNavigationTarget(pageImage);
            currentPage!.Images.Add(pageImage);
            pageImage.InlineDrawToken = AllocateInlineImageDrawToken(currentPage);
            sb.Append(pageImage.InlineDrawToken);

            pageDirty = true;
        }

        private static bool ChildImagesOwnDrawingAccessibility(OfficeDrawing drawing, PdfDrawingStyle style) {
            if (style.Decorative || !string.IsNullOrWhiteSpace(style.AlternativeText) || drawing.Elements.Count == 0) return false;
            for (int i = 0; i < drawing.Elements.Count; i++) {
                OfficeDrawingElement element = drawing.Elements[i];
                if (element is OfficeDrawingImage image) {
                    if (string.IsNullOrWhiteSpace(image.AlternativeText)) return false;
                } else if (element is OfficeDrawingGroup group) {
                    if (!ChildImagesOwnDrawingAccessibility(group.InnerDrawing, style)) return false;
                } else if (element is OfficeDrawingEffectGroup effectGroup) {
                    if (!ChildImagesOwnDrawingAccessibility(effectGroup.InnerDrawing, style)) return false;
                } else {
                    return false;
                }
            }

            return true;
        }

        private int? AppendDrawingMarkedContentBegin(PdfDrawingStyle style, out bool markedContent, bool recordDrawingEvidence = true, long? logicalOrder = null) {
            EnsurePage();

            if (_suppressCanvasAccessibilityWrappers) {
                markedContent = false;
                return null;
            }

            if (recordDrawingEvidence) {
                currentPage!.Drawings.Add(new PdfGeneratedDrawingAccessibilityEvidence(!string.IsNullOrWhiteSpace(style.AlternativeText), style.Decorative));
            }

            if (style.Decorative) {
                AppendArtifactBegin(sb, emitGeneratedStructure);
                markedContent = emitGeneratedStructure;
                return null;
            }

            if (string.IsNullOrWhiteSpace(style.AlternativeText)) {
                markedContent = false;
                return null;
            }

            int? markedContentId = RegisterFigureStructureElement(style.AlternativeText!, _canvasStructureParentElement, logicalOrder);
            int? structElementIndex = FindStructElementIndex(currentPage, markedContentId, "Figure");
            sb.Append("/Figure << /Alt ")
                .Append(PdfSyntaxEscaper.TextString(style.AlternativeText!));
            if (markedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }

            sb.Append(" >> BDC\n");
            markedContent = true;
            return structElementIndex;
        }

        private void AppendDrawingMarkedContentEnd(bool markedContent) {
            if (markedContent) {
                sb.Append("EMC\n");
            }
        }

        private void RenderShapeBlock(ShapeBlock block, double containerX, double containerWidth) {
            double frameMarginLeft = currentOpts.MarginLeft;
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Shape");
            double spacingBefore = PlaceFixedFlowBlock("Shape", block.Shape.Width, block.Shape.Height,
                style.SpacingBefore, style.SpacingAfter, ref containerWidth);
            if (spacingBefore > 0) y -= spacingBefore;
            RecordFlowPlacement(y);
            containerX += currentOpts.MarginLeft - frameMarginLeft;
            int? structElementIndex = DrawShapeAt(block, style, containerX, containerWidth, y);
            AddShapeLinkAnnotation(block, style, containerX, containerWidth, y, structElementIndex);
            DrawDebugFlowObjectBox(GetAlignedObjectX(containerX, containerWidth, block.Shape.Width, style.Align), y - block.Shape.Height, block.Shape.Width, block.Shape.Height);
            y -= block.Shape.Height + style.SpacingAfter;
        }

        private void RenderDrawingBlock(DrawingBlock block, double containerX, double containerWidth) {
            double frameMarginLeft = currentOpts.MarginLeft;
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Drawing");
            var box = PlaceDrawingFlowBlock(block, style, ref containerWidth);
            if (box.SpacingBefore > 0) y -= box.SpacingBefore;
            RecordFlowPlacement(y);
            containerX += currentOpts.MarginLeft - frameMarginLeft;
            int? structElementIndex = DrawDrawingAt(block, style, containerX, containerWidth, y);
            AddDrawingLinkAnnotation(block, style, containerX, containerWidth, y, structElementIndex);
            DrawDebugFlowObjectBox(GetAlignedObjectX(containerX, containerWidth, box.Width, style.Align), y - box.Height, box.Width, box.Height);
            y -= box.Height + style.SpacingAfter;
        }

        private void KeepFixedBlockWithNext(double objectWidth, double objectHeight, double spacingBefore, double spacingAfter,
            IList<IPdfBlock> blocks, int blockIndex) {
            while (true) {
                double needed = ResolveTopLevelSpacingBefore(spacingBefore) + objectHeight + spacingAfter;
                EnsureFixedFlowBlockFits("Kept object", objectWidth, objectHeight + spacingAfter, GetMaximumFixedFlowWidth(width));
                if (objectWidth > width + .001D || ShouldAdvanceForBlockHeight(needed)) {
                    NewBlockFrame();
                    continue;
                }
                double nextHeight = MeasureKeepWithNextChainHeight(blocks, blockIndex + 1, currentOpts.MarginLeft,
                    width, currentOpts.DefaultFontSize, needed);
                double keepHeight = needed + nextHeight;
                if (nextHeight <= .001D || keepHeight > GetMaximumBlockContinuationHeight() + .001D || !ShouldAdvanceForBlockHeight(keepHeight)) break;
                NewBlockFrame();
            }
        }

        private static double GetAlignedObjectX(double containerX, double containerWidth, double objectWidth, PdfAlign align) {
            if (align == PdfAlign.Center) return containerX + Math.Max(0, (containerWidth - objectWidth) / 2);
            if (align == PdfAlign.Right) return containerX + Math.Max(0, containerWidth - objectWidth);
            return containerX;
        }

        private void AddShapeLinkAnnotation(ShapeBlock shape, PdfDrawingStyle style, double containerX, double containerWidth, double topY, int? structElementIndex = null) {
            if (string.IsNullOrEmpty(shape.LinkUri)) {
                return;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, shape.Shape.Width, style.Align);
            var canvas = (shape.Shape.Transform ?? OfficeTransform.Identity)
                .TransformRectangleBounds(0, 0, shape.Shape.Width, shape.Shape.Height);
            double x1 = x + canvas.Left;
            double x2 = x + canvas.Right;
            double y1 = topY - canvas.Bottom;
            double y2 = topY - canvas.Top;
            double halfStroke = Math.Max(0.5D, Math.Max(0D, shape.Shape.StrokeWidth) / 2D);
            if (x2 <= x1) {
                x1 -= halfStroke;
                x2 += halfStroke;
            }

            if (y2 <= y1) {
                y1 -= halfStroke;
                y2 += halfStroke;
            }

            if (TryGetShapePaintBounds(shape.Shape, x, topY - shape.Shape.Height, null, "Body", "shape",
                    out double paintX, out double paintY, out double paintWidth, out double paintHeight, conservativeWhenUnmeasured: true)) {
                x1 = Math.Min(x1, paintX); x2 = Math.Max(x2, paintX + paintWidth);
                y1 = Math.Min(y1, paintY); y2 = Math.Max(y2, paintY + paintHeight);
            }

            foreach (OfficeShape marker in CreateLineMarkerShapes(shape.Shape)) {
                if (!TryGetShapePaintBounds(marker, x, topY - marker.Height, null, "Body", "marker",
                        out double markerX, out double markerY, out double markerWidth, out double markerHeight, conservativeWhenUnmeasured: true)) continue;
                x1 = Math.Min(x1, markerX); x2 = Math.Max(x2, markerX + markerWidth);
                y1 = Math.Min(y1, markerY); y2 = Math.Max(y2, markerY + markerHeight);
            }

            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = shape.LinkUri!, Contents = shape.LinkContents, StructElementIndex = structElementIndex });
        }

        private void AddDrawingLinkAnnotation(DrawingBlock drawing, PdfDrawingStyle style, double containerX, double containerWidth, double topY, int? structElementIndex = null) {
            if (string.IsNullOrEmpty(drawing.LinkUri)) {
                return;
            }

            var box = ResolveDrawingFlowBox(drawing, style, containerWidth);
            double x = GetAlignedObjectX(containerX, containerWidth, box.Width, style.Align);
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - box.Height, X2 = x + box.Width, Y2 = topY, Uri = drawing.LinkUri!, Contents = drawing.LinkContents, StructElementIndex = structElementIndex });
        }

    }
}
