using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static PageImage CreatePageImage(ImageBlock block, PdfImageStyle style, double targetX, double targetBottomY) {
        double drawX = targetX;
        double drawY = targetBottomY;
        double drawWidth = block.Width;
        double drawHeight = block.Height;
        OfficeClipPath? clipPath = style.ClipPath?.Clone();

        if (style.Fit != OfficeImageFit.Stretch) {
            double imageAspect = block.Info.Width / (double)block.Info.Height;
            double targetAspect = block.Width / block.Height;

            if (style.Fit == OfficeImageFit.Contain) {
                if (targetAspect > imageAspect) {
                    drawHeight = block.Height;
                    drawWidth = drawHeight * imageAspect;
                    drawX = targetX + (block.Width - drawWidth) / 2D;
                } else {
                    drawWidth = block.Width;
                    drawHeight = drawWidth / imageAspect;
                    drawY = targetBottomY + (block.Height - drawHeight) / 2D;
                }
            } else {
                if (targetAspect > imageAspect) {
                    drawWidth = block.Width;
                    drawHeight = drawWidth / imageAspect;
                    drawY = targetBottomY + (block.Height - drawHeight) / 2D;
                } else {
                    drawHeight = block.Height;
                    drawWidth = drawHeight * imageAspect;
                    drawX = targetX + (block.Width - drawWidth) / 2D;
                }

                if (clipPath == null) {
                    clipPath = OfficeClipPath.Rectangle(block.Width, block.Height);
                }
            }
        }

        return new PageImage {
            Data = block.Data,
            Info = block.Info,
            X = drawX,
            Y = drawY,
            W = drawWidth,
            H = drawHeight,
            ClipPath = clipPath,
            ClipX = targetX,
            ClipY = targetBottomY,
            ClipHeight = block.Height,
            AlternativeText = style.AlternativeText
        };
    }

    private static void AddHeaderFooterImages(LayoutResult.Page page, PdfOptions options, int variantPageNumber) {
        foreach (PdfHeaderFooterImage image in options.GetHeaderImagesForPage(variantPageNumber)) {
            AddHeaderFooterImage(page, options, image, isHeader: true);
        }

        foreach (PdfHeaderFooterImage image in options.GetFooterImagesForPage(variantPageNumber)) {
            AddHeaderFooterImage(page, options, image, isHeader: false);
        }
    }

    private static void AddHeaderFooterImage(LayoutResult.Page page, PdfOptions options, PdfHeaderFooterImage image, bool isHeader) {
        double contentLeft = options.MarginLeft;
        double contentWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        if (image.Width > contentWidth + 0.001D) {
            throw new ArgumentException("PDF " + (isHeader ? "header" : "footer") + " image must fit inside the page content width.");
        }

        double x = contentLeft;
        if (image.Align == PdfAlign.Center) {
            x = contentLeft + Math.Max(0D, (contentWidth - image.Width) / 2D);
        } else if (image.Align == PdfAlign.Right) {
            x = contentLeft + Math.Max(0D, contentWidth - image.Width);
        }

        double y = isHeader
            ? options.PageHeight - options.MarginTop + options.HeaderOffsetY - image.Height
            : options.MarginBottom - options.FooterOffsetY;
        if (y < -0.001D || y + image.Height > options.PageHeight + 0.001D) {
            throw new ArgumentException("PDF " + (isHeader ? "header" : "footer") + " image must fit inside the page bounds.");
        }

        ImageBlock block = image.ToImageBlock();
        PageImage pageImage = CreatePageImage(block, block.Style ?? new PdfImageStyle(), x, y);
        pageImage.IsBackgroundDecoration = string.IsNullOrWhiteSpace(pageImage.AlternativeText);
        page.Images.Add(pageImage);
    }

    private static string BuildHeaderFooterShapes(LayoutResult.Page page, PdfOptions options, int variantPageNumber) {
        var sb = new StringBuilder();
        foreach (PdfHeaderFooterShape shape in options.GetHeaderShapesForPage(variantPageNumber)) {
            AddHeaderFooterShape(sb, page, options, shape, isHeader: true);
        }

        foreach (PdfHeaderFooterShape shape in options.GetFooterShapesForPage(variantPageNumber)) {
            AddHeaderFooterShape(sb, page, options, shape, isHeader: false);
        }

        return sb.ToString();
    }

    private static void AddHeaderFooterShape(StringBuilder sb, LayoutResult.Page page, PdfOptions options, PdfHeaderFooterShape headerFooterShape, bool isHeader) {
        ShapeBlock block = headerFooterShape.ToShapeBlock();
        PdfDrawingStyle style = block.Style ?? new PdfDrawingStyle();
        PdfDocument.ValidateDrawingStyle(style, isHeader ? "Header shape" : "Footer shape");

        double contentLeft = options.MarginLeft;
        double contentWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        if (block.Shape.Width > contentWidth + 0.001D) {
            throw new ArgumentException("PDF " + (isHeader ? "header" : "footer") + " shape must fit inside the page content width.");
        }

        double x = GetHeaderFooterAlignedObjectX(contentLeft, contentWidth, block.Shape.Width, style.Align);
        double bottomY = isHeader
            ? options.PageHeight - options.MarginTop + options.HeaderOffsetY - block.Shape.Height
            : options.MarginBottom - options.FooterOffsetY;
        if (bottomY < -0.001D || bottomY + block.Shape.Height > options.PageHeight + 0.001D) {
            throw new ArgumentException("PDF " + (isHeader ? "header" : "footer") + " shape must fit inside the page bounds.");
        }

        DrawHeaderFooterShapeGeometryAt(sb, page, block.Shape, x, bottomY);
    }

    private static double GetHeaderFooterAlignedObjectX(double containerX, double containerWidth, double objectWidth, PdfAlign align) {
        if (align == PdfAlign.Center) return containerX + Math.Max(0, (containerWidth - objectWidth) / 2);
        if (align == PdfAlign.Right) return containerX + Math.Max(0, containerWidth - objectWidth);
        return containerX;
    }

    private static PdfColor? ToHeaderFooterPdfColor(OfficeColor? color) =>
        color.HasValue ? PdfColor.FromOfficeColorOrNull(color.Value) : null;

    private static string? EnsureHeaderFooterGraphicsState(LayoutResult.Page page, double fillOpacity, double strokeOpacity) {
        if (fillOpacity >= 1D && strokeOpacity >= 1D) {
            return null;
        }

        for (int i = 0; i < page.GraphicsStates.Count; i++) {
            var existing = page.GraphicsStates[i];
            if (existing.FillOpacity.Equals(fillOpacity) && existing.StrokeOpacity.Equals(strokeOpacity)) {
                return existing.Name;
            }
        }

        string name = "GS" + (page.GraphicsStates.Count + 1).ToString(CultureInfo.InvariantCulture);
        page.GraphicsStates.Add(new PageGraphicsState {
            Name = name,
            FillOpacity = fillOpacity,
            StrokeOpacity = strokeOpacity
        });
        return name;
    }

    private static string? EnsureHeaderFooterOpacityState(LayoutResult.Page page, OfficeShape shape) {
        bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null) && shape.Kind != OfficeShapeKind.Line;
        bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
        double fillOpacity = hasFill ? shape.FillOpacity ?? 1D : 1D;
        double strokeOpacity = hasStroke ? shape.StrokeOpacity ?? 1D : 1D;
        return EnsureHeaderFooterGraphicsState(page, fillOpacity, strokeOpacity);
    }

    private static string? EnsureHeaderFooterLinearGradient(LayoutResult.Page page, OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
        var gradient = shape.FillGradient;
        if (gradient == null || shape.Kind == OfficeShapeKind.Line) {
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

        for (int i = 0; i < page.Shadings.Count; i++) {
            var existing = page.Shadings[i];
            if (existing.StartColor.Equals(start) &&
                existing.EndColor.Equals(end) &&
                existing.X0.Equals(x0) &&
                existing.Y0.Equals(y0) &&
                existing.X1.Equals(x1) &&
                existing.Y1.Equals(y1)) {
                return existing.Name;
            }
        }

        string name = "SH" + (page.Shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
        page.Shadings.Add(new PageShading {
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

    private static void DrawHeaderFooterShapeShadowAt(StringBuilder sb, LayoutResult.Page page, OfficeShape shape, double xShape, double bottomY) {
        var shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D) {
            return;
        }

        PdfColor shadowColor = PdfColor.FromOfficeColor(shadow.Color);
        double shadowX = xShape + shadow.OffsetX;
        double shadowBottomY = bottomY - shadow.OffsetY;
        string? shadowState = EnsureHeaderFooterGraphicsState(page, shadow.Opacity, shadow.Opacity);

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (shadowState != null) {
            content.GraphicsState(shadowState);
        }

        if (shape.Transform.HasValue) {
            DrawTransformedShape(
                sb,
                shape,
                shape.Kind == OfficeShapeKind.Line ? null : shadowColor,
                shape.Kind == OfficeShapeKind.Line ? shadowColor : null,
                null,
                shadowX,
                shadowBottomY);
        } else if (shape.Kind == OfficeShapeKind.Line) {
            DrawLine(sb, shadowColor, shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
        } else if (shape.Kind == OfficeShapeKind.RoundedRectangle) {
            DrawRoundedRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height, shape.CornerRadius);
        } else if (shape.Kind == OfficeShapeKind.Rectangle) {
            DrawRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
        } else if (shape.Kind == OfficeShapeKind.Ellipse) {
            DrawEllipse(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
        } else if (shape.Kind == OfficeShapeKind.Polygon) {
            DrawPolygon(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
        } else if (shape.Kind == OfficeShapeKind.Path) {
            DrawPath(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, shadowX, shadowBottomY, shape.Height);
        }

        content.RestoreState();
    }

    private static void DrawHeaderFooterShapeGeometryAt(StringBuilder sb, LayoutResult.Page page, OfficeShape shape, double xShape, double bottomY) {
        DrawHeaderFooterShapeShadowAt(sb, page, shape, xShape, bottomY);

        string? opacityState = EnsureHeaderFooterOpacityState(page, shape);
        if (opacityState != null) {
            new ContentStreamBuilder(sb)
                .SaveState()
                .GraphicsState(opacityState);
        }

        if (shape.Transform.HasValue) {
            string? shadingName = EnsureHeaderFooterLinearGradient(page, shape, xShape, bottomY, localCoordinates: true);
            DrawTransformedShape(sb, shape, shadingName == null ? ToHeaderFooterPdfColor(shape.FillColor) : null, ToHeaderFooterPdfColor(shape.StrokeColor), shadingName, xShape, bottomY);
        } else {
            if (shape.ClipPath != null) {
                new ContentStreamBuilder(sb)
                    .SaveState();
                AppendClipPath(sb, shape.ClipPath, xShape, bottomY, shape.Height);
            }

            string? shadingName = EnsureHeaderFooterLinearGradient(page, shape, xShape, bottomY, localCoordinates: false);
            if (shadingName != null) {
                DrawGradientShape(sb, shape, shadingName, xShape, bottomY);
            }

            PdfColor? fillColor = shadingName == null ? ToHeaderFooterPdfColor(shape.FillColor) : null;
            if (shape.Kind == OfficeShapeKind.Line) {
                DrawLine(sb, ToHeaderFooterPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
            } else if (shape.Kind == OfficeShapeKind.RoundedRectangle) {
                DrawRoundedRectangle(sb, fillColor, ToHeaderFooterPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height, shape.CornerRadius);
            } else if (shape.Kind == OfficeShapeKind.Rectangle) {
                DrawRectangle(sb, fillColor, ToHeaderFooterPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeShapeKind.Ellipse) {
                DrawEllipse(sb, fillColor, ToHeaderFooterPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeShapeKind.Polygon) {
                DrawPolygon(sb, fillColor, ToHeaderFooterPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
            } else if (shape.Kind == OfficeShapeKind.Path) {
                DrawPath(sb, fillColor, ToHeaderFooterPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, xShape, bottomY, shape.Height);
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

}
