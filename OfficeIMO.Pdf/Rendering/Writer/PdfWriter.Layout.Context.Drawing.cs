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
            bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null || shape.FillRadialGradient != null) && shape.Kind != OfficeIMO.Drawing.OfficeShapeKind.Line;
            bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
            double fillOpacity = hasFill ? shape.FillOpacity ?? 1D : 1D;
            double strokeOpacity = hasStroke ? shape.StrokeOpacity ?? 1D : 1D;
            return EnsureGraphicsState(fillOpacity, strokeOpacity);
        }

        private string? EnsureFillGradient(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
            if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) return null;
            EnsurePage();
            if (shape.FillRadialGradient != null) return EnsureRadialShading(currentPage!.Shadings, shape.FillRadialGradient);
            var gradient = shape.FillGradient;
            if (gradient == null) {
                return null;
            }

            double originX = localCoordinates ? 0D : xShape;
            double originY = localCoordinates ? 0D : bottomY;
            double x0 = originX + gradient.StartX * shape.Width;
            double y0 = originY + shape.Height - gradient.StartY * shape.Height;
            double x1 = originX + gradient.EndX * shape.Width;
            double y1 = originY + shape.Height - gradient.EndY * shape.Height;
            return EnsureAxialShading(currentPage!.Shadings, gradient, x0, y0, x1, y1);
        }

        private void DrawShapeShadowAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            var shadow = shape.Shadow;
            if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) return;
            double coreOpacity = shadow.Opacity * shadow.Color.A / 255D;
            PdfColor shadowColor = PdfColor.FromRgb(shadow.Color.R, shadow.Color.G, shadow.Color.B);
            double shadowX = xShape + shadow.OffsetX;
            double shadowBottomY = bottomY - shadow.OffsetY;
            ResolveShadowGeometry(shape, out bool hasFill, out bool hasStroke);
            if (shadow.BlurRadius > 0D) {
                const int layers = 4;
                for (int index = layers; index >= 1; index--) {
                    double factor = index / (double)layers;
                    double opacity = coreOpacity * (0.04D + (layers - index + 1) * 0.05D);
                    DrawShapeShadowLayerAt(
                        shape,
                        shadowColor,
                        shadowX,
                        shadowBottomY,
                        Math.Max(1D, Math.Max(0D, shape.StrokeWidth) + shadow.BlurRadius * 2D * factor),
                        opacity,
                        hasFill,
                        hasStroke: true);
                }
            }
            DrawShapeShadowLayerAt(shape, shadowColor, shadowX, shadowBottomY, Math.Max(0D, shape.StrokeWidth), coreOpacity, hasFill, hasStroke);
            pageDirty = true;
        }

        private void DrawShapeShadowLayerAt(
            OfficeShape shape,
            PdfColor color,
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
            DrawShapeShadowAt(shape, xShape, bottomY);

            string? opacityState = EnsureOpacityState(shape);
            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .GraphicsState(opacityState);
            }

            if (shape.Transform.HasValue) {
                pageDirty = true;
                string? shadingName = EnsureFillGradient(shape, xShape, bottomY, localCoordinates: true);
                DrawTransformedShape(sb, shape, shadingName == null ? ToPdfColor(shape.FillColor) : null, ToPdfColor(shape.StrokeColor), shadingName, xShape, bottomY);
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
            for (int i = 0; i < block.Drawing.Elements.Count; i++) {
                if (block.Drawing.Elements[i] is OfficeDrawingShape shape) {
                    double xShape = xDrawing + shape.X;
                    double bottomY = topY - shape.Y - shape.Shape.Height;
                    DrawShapeGeometryAt(shape.Shape, xShape, bottomY);
                } else if (block.Drawing.Elements[i] is OfficeDrawingText text) {
                    DrawDrawingTextAt(text, xDrawing + text.X, topY - text.Y);
                }
            }

            AppendDrawingMarkedContentEnd(markedContent);
            return structElementIndex;
        }

        private void DrawDrawingTextAt(OfficeDrawingText text, double x, double topY) {
            if (string.IsNullOrEmpty(text.Text)) {
                return;
            }

            PdfStandardFont baseFont = ResolveDrawingTextFont(text.Font);
            double size = text.Font.Size;
            double leading = text.LineHeight ?? size * 1.2D;
            PdfColor? color = ToPdfColor(text.Color);
            var runs = new[] {
                new TextRun(
                    text.Text,
                    bold: text.Font.IsBold,
                    underline: text.Font.IsUnderline,
                    color: color,
                    italic: text.Font.IsItalic,
                    fontSize: size,
                    font: baseFont,
                    fontFamily: text.Font.FamilyName)
            };
            var block = new RichParagraphBlock(runs, MapDrawingTextAlignment(text.Alignment), color);
            var wrap = WrapRichRunsCore(runs, text.Width, size, baseFont, leading, null, DefaultParagraphTabStopWidth, currentOpts);
            if (wrap.Lines.Count == 0) {
                return;
            }

            WriteClippedRichParagraph(
                sb,
                block,
                wrap.Lines,
                wrap.LineHeights,
                currentOpts,
                FirstTextBaselineFromTop(baseFont, size, topY),
                size,
                leading,
                currentPage!.Annotations,
                x,
                topY - text.Height,
                text.Width,
                text.Height,
                x,
                text.Width,
                structureType: null,
                markedContentId: null,
                structurePage: null);
            MarkRichFonts(runs);
            pageDirty = true;
        }

        private PdfStandardFont ResolveDrawingTextFont(OfficeFontInfo font) {
            if (!string.IsNullOrWhiteSpace(font.FamilyName) && PdfStandardFontMapper.TryMapFontFamily(font.FamilyName, out PdfStandardFont mapped)) {
                return ChooseNormal(mapped);
            }

            return ChooseNormal(currentOpts.DefaultFont);
        }

        private static PdfAlign MapDrawingTextAlignment(OfficeTextAlignment alignment) {
            if (alignment == OfficeTextAlignment.Center) {
                return PdfAlign.Center;
            }

            if (alignment == OfficeTextAlignment.Right) {
                return PdfAlign.Right;
            }

            return PdfAlign.Left;
        }

        private int? AppendDrawingMarkedContentBegin(PdfDrawingStyle style, out bool markedContent) {
            EnsurePage();

            if (_suppressCanvasAccessibilityWrappers) {
                markedContent = false;
                return null;
            }

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

            int? markedContentId = RegisterFigureStructureElement(style.AlternativeText!, _canvasStructureParentElementIndex);
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
            DrawDebugFlowObjectBox(GetAlignedObjectX(containerX, containerWidth, block.Shape.Width, style.Align), y - block.Shape.Height, block.Shape.Width, block.Shape.Height);
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
            DrawDebugFlowObjectBox(GetAlignedObjectX(containerX, containerWidth, block.Drawing.Width, style.Align), y - block.Drawing.Height, block.Drawing.Width, block.Drawing.Height);
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

