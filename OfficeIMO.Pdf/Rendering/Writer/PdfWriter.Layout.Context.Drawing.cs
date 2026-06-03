using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static PdfColor? ToPdfColor(OfficeIMO.Drawing.OfficeColor? color) =>
            color.HasValue ? PdfColor.FromOfficeColorOrNull(color.Value) : null;

        private string? EnsureGraphicsState(double fillOpacity, double strokeOpacity) {
            if (fillOpacity >= 1D && strokeOpacity >= 1D) {
                return null;
            }

            EnsurePage();
            for (int i = 0; i < currentPage!.GraphicsStates.Count; i++) {
                var existing = currentPage.GraphicsStates[i];
                if (existing.FillOpacity.Equals(fillOpacity) && existing.StrokeOpacity.Equals(strokeOpacity)) {
                    return existing.Name;
                }
            }

            string name = "GS" + (currentPage.GraphicsStates.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.GraphicsStates.Add(new PageGraphicsState {
                Name = name,
                FillOpacity = fillOpacity,
                StrokeOpacity = strokeOpacity
            });
            return name;
        }

        private string? EnsureOpacityState(OfficeIMO.Drawing.OfficeShape shape) {
            bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null) && shape.Kind != OfficeIMO.Drawing.OfficeShapeKind.Line;
            bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
            double fillOpacity = hasFill ? shape.FillOpacity ?? 1D : 1D;
            double strokeOpacity = hasStroke ? shape.StrokeOpacity ?? 1D : 1D;
            return EnsureGraphicsState(fillOpacity, strokeOpacity);
        }

        private string? EnsureLinearGradient(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
            var gradient = shape.FillGradient;
            if (gradient == null || shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                return null;
            }

            var start = gradient.Stops[0].Color;
            var end = gradient.Stops[1].Color;
            double originX = localCoordinates ? 0D : xShape;
            double originY = localCoordinates ? 0D : bottomY;
            double x0 = originX + gradient.StartX * shape.Width;
            double y0 = originY + shape.Height - gradient.StartY * shape.Height;
            double x1 = originX + gradient.EndX * shape.Width;
            double y1 = originY + shape.Height - gradient.EndY * shape.Height;

            EnsurePage();
            for (int i = 0; i < currentPage!.Shadings.Count; i++) {
                var existing = currentPage.Shadings[i];
                if (existing.StartColor.Equals(start) &&
                    existing.EndColor.Equals(end) &&
                    existing.X0.Equals(x0) &&
                    existing.Y0.Equals(y0) &&
                    existing.X1.Equals(x1) &&
                    existing.Y1.Equals(y1)) {
                    return existing.Name;
                }
            }

            string name = "SH" + (currentPage.Shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.Shadings.Add(new PageShading {
                Name = name,
                StartColor = start,
                EndColor = end,
                X0 = x0,
                Y0 = y0,
                X1 = x1,
                Y1 = y1
            });
            return name;
        }

        private void DrawShapeShadowAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            var shadow = shape.Shadow;
            if (shadow == null || shadow.Opacity <= 0D) {
                return;
            }

            PdfColor shadowColor = PdfColor.FromOfficeColor(shadow.Color);
            double shadowX = xShape + shadow.OffsetX;
            double shadowBottomY = bottomY - shadow.OffsetY;
            string? shadowState = EnsureGraphicsState(shadow.Opacity, shadow.Opacity);

            var content = new ContentStreamBuilder(sb)
                .SaveState();
            if (shadowState != null) {
                content.GraphicsState(shadowState);
            }

            if (shape.Transform.HasValue) {
                DrawTransformedShape(
                    sb,
                    shape,
                    shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line ? null : shadowColor,
                    shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line ? shadowColor : null,
                    null,
                    shadowX,
                    shadowBottomY);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                DrawLine(sb, shadowColor, shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle) {
                DrawRoundedRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height, shape.CornerRadius);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle) {
                DrawRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Ellipse) {
                DrawEllipse(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Polygon) {
                DrawPolygon(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                DrawPath(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, shadowX, shadowBottomY, shape.Height);
            }

            content.RestoreState();
            pageDirty = true;
        }

        private void DrawShapeGeometryAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            DrawShapeShadowAt(shape, xShape, bottomY);

            string? opacityState = EnsureOpacityState(shape);
            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .GraphicsState(opacityState);
            }

            if (shape.Transform.HasValue) {
                pageDirty = true;
                string? shadingName = EnsureLinearGradient(shape, xShape, bottomY, localCoordinates: true);
                DrawTransformedShape(sb, shape, shadingName == null ? ToPdfColor(shape.FillColor) : null, ToPdfColor(shape.StrokeColor), shadingName, xShape, bottomY);
            } else {
                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .SaveState();
                    AppendClipPath(sb, shape.ClipPath, xShape, bottomY, shape.Height);
                }

                string? shadingName = EnsureLinearGradient(shape, xShape, bottomY, localCoordinates: false);
                if (shadingName != null) {
                    pageDirty = true;
                    DrawGradientShape(sb, shape, shadingName, xShape, bottomY);
                }

                PdfColor? fillColor = shadingName == null ? ToPdfColor(shape.FillColor) : null;
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
                    DrawPolygon(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                    pageDirty = true;
                    DrawPath(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, xShape, bottomY, shape.Height);
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
            double xDrawing = GetAlignedObjectX(containerX, containerWidth, block.Drawing.Width, style.Align);
            bool markedContent;
            int? structElementIndex = AppendDrawingMarkedContentBegin(style, out markedContent);
            for (int i = 0; i < block.Drawing.Shapes.Count; i++) {
                var item = block.Drawing.Shapes[i];
                double xShape = xDrawing + item.X;
                double bottomY = topY - item.Y - item.Shape.Height;
                DrawShapeGeometryAt(item.Shape, xShape, bottomY);
            }

            AppendDrawingMarkedContentEnd(markedContent);
            return structElementIndex;
        }

        private int? AppendDrawingMarkedContentBegin(PdfDrawingStyle style, out bool markedContent) {
            EnsurePage();
            currentPage!.Drawings.Add(new PdfGeneratedDrawingAccessibilityEvidence(!string.IsNullOrWhiteSpace(style.AlternativeText), style.Decorative));

            if (style.Decorative) {
                AppendArtifactBegin(sb, emitGeneratedStructure);
                markedContent = emitGeneratedStructure;
                return null;
            }

            if (string.IsNullOrWhiteSpace(style.AlternativeText)) {
                markedContent = false;
                return null;
            }

            int? markedContentId = RegisterFigureStructureElement(style.AlternativeText!);
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
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Shape");
            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double needed = spacingBefore + block.Shape.Height + style.SpacingAfter;
            EnsureFixedFlowBlockFits("Shape", block.Shape.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            int? structElementIndex = DrawShapeAt(block, style, containerX, containerWidth, y);
            AddShapeLinkAnnotation(block, style, containerX, containerWidth, y, structElementIndex);
            y -= block.Shape.Height + style.SpacingAfter;
        }

        private void RenderDrawingBlock(DrawingBlock block, double containerX, double containerWidth) {
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Drawing");
            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double needed = spacingBefore + block.Drawing.Height + style.SpacingAfter;
            EnsureFixedFlowBlockFits("Drawing", block.Drawing.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            int? structElementIndex = DrawDrawingAt(block, style, containerX, containerWidth, y);
            AddDrawingLinkAnnotation(block, style, containerX, containerWidth, y, structElementIndex);
            y -= block.Drawing.Height + style.SpacingAfter;
        }

        private void KeepFixedBlockWithNext(double needed, double nextHeight) {
            double keepHeight = needed + nextHeight;
            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                NewPage();
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
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - shape.Shape.Height, X2 = x + shape.Shape.Width, Y2 = topY, Uri = shape.LinkUri!, Contents = shape.LinkContents, StructElementIndex = structElementIndex });
        }

        private void AddDrawingLinkAnnotation(DrawingBlock drawing, PdfDrawingStyle style, double containerX, double containerWidth, double topY, int? structElementIndex = null) {
            if (string.IsNullOrEmpty(drawing.LinkUri)) {
                return;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, drawing.Drawing.Width, style.Align);
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - drawing.Drawing.Height, X2 = x + drawing.Drawing.Width, Y2 = topY, Uri = drawing.LinkUri!, Contents = drawing.LinkContents, StructElementIndex = structElementIndex });
        }

    }
}


