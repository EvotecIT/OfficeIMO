using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double HeaderFooterInlineGap = 4D;

    private static PageImage CreatePageImage(ImageBlock block, PdfImageStyle style, double targetX, double targetBottomY) =>
        CreatePageImage(block, style, targetX, targetBottomY, block.Width, block.Height);

    private static PageImage CreatePageImage(ImageBlock block, PdfImageStyle style, double targetX, double targetBottomY, double targetWidth, double targetHeight) {
        OfficeClipPath? clipPath = style.ClipPath?.Scale(targetWidth / block.Width, targetHeight / block.Height);
        PdfImageSourceCrop? sourceCrop = style.SourceCrop;
        OfficeImageSourceCrop crop = sourceCrop?.ToOfficeImageSourceCrop() ?? default;
        OfficeImageRenderPlan renderPlan = OfficeImageRenderPlan.CreateBottomLeft(
            block.Info.Width,
            block.Info.Height,
            targetX,
            targetBottomY,
            targetWidth,
            targetHeight,
            style.Fit,
            crop);
        if (renderPlan.RequiresTargetClip && clipPath == null) {
            clipPath = OfficeClipPath.Rectangle(targetWidth, targetHeight);
        }

        return new PageImage {
            Data = block.Data,
            Info = block.Info,
            X = renderPlan.ImagePlacement.X,
            Y = renderPlan.ImagePlacement.Y,
            W = renderPlan.ImagePlacement.Width,
            H = renderPlan.ImagePlacement.Height,
            ClipPath = clipPath,
            ClipX = targetX,
            ClipY = targetBottomY,
            ClipHeight = targetHeight,
            SourceCrop = sourceCrop?.Clone(),
            RotationAngle = style.RotationAngle,
            AlternativeText = style.AlternativeText
        };
    }

    private static void GetImageAnnotationBounds(PdfImageStyle style, PageImage pageImage, double targetX, double targetBottomY, double targetWidth, double targetHeight, out double x1, out double y1, out double x2, out double y2) {
        x1 = pageImage.X;
        y1 = pageImage.Y;
        x2 = pageImage.X + pageImage.W;
        y2 = pageImage.Y + pageImage.H;

        if (style.Fit != OfficeImageFit.Cover && style.ClipPath == null && style.SourceCrop?.HasCrop != true) {
            return;
        }

        x1 = targetX;
        y1 = targetBottomY;
        x2 = targetX + targetWidth;
        y2 = targetBottomY + targetHeight;
    }

    private static void AddHeaderFooterImages(
        LayoutResult.Page page,
        PdfOptions options,
        int variantPageNumber,
        int pageNumber,
        int totalPages,
        int documentPages) {
        AddHeaderFooterImages(
            page,
            options,
            options.GetHeaderImagesForPage(variantPageNumber),
            variantPageNumber,
            pageNumber,
            totalPages,
            documentPages,
            isHeader: true);
        AddHeaderFooterImages(
            page,
            options,
            options.GetFooterImagesForPage(variantPageNumber),
            variantPageNumber,
            pageNumber,
            totalPages,
            documentPages,
            isHeader: false);
    }

    private static void AddHeaderFooterImages(
        LayoutResult.Page page,
        PdfOptions options,
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> images,
        int variantPageNumber,
        int pageNumber,
        int totalPages,
        int documentPages,
        bool isHeader) {
        var consumedWidths = new System.Collections.Generic.Dictionary<PdfAlign, double>();
        foreach (PdfHeaderFooterImage image in images) {
            double textWidth = MeasureHeaderFooterTextWidth(
                options,
                variantPageNumber,
                pageNumber,
                totalPages,
                documentPages,
                image.Align,
                isHeader);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, image.Align);
            double groupWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth);
            double groupX = AlignHeaderFooterGroup(options, groupWidth, image.Align);
            double consumedWidth = consumedWidths.TryGetValue(image.Align, out double value) ? value : 0D;
            double imageX = groupX +
                (textWidth > 0D ? textWidth + HeaderFooterInlineGap : 0D) +
                consumedWidth;

            AddHeaderFooterImage(page, options, image, imageX, isHeader);
            consumedWidths[image.Align] = consumedWidth + image.Width + HeaderFooterInlineGap;
        }
    }

    private static void AddHeaderFooterImage(LayoutResult.Page page, PdfOptions options, PdfHeaderFooterImage image, double x, bool isHeader) {
        double contentLeft = options.MarginLeft;
        double contentWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        if (image.Width > contentWidth + 0.001D) {
            throw new ArgumentException("PDF " + (isHeader ? "header" : "footer") + " image must fit inside the page content width.");
        }

        if (x < contentLeft - 0.001D || x + image.Width > contentLeft + contentWidth + 0.001D) {
            throw new ArgumentException("Combined PDF " + (isHeader ? "header" : "footer") + " text and images must fit inside the page content width.");
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

    private static double MeasureHeaderFooterImagesWidth(System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> images, PdfAlign align) {
        double width = 0D;
        int count = 0;
        foreach (PdfHeaderFooterImage image in images) {
            if (image.Align != align) {
                continue;
            }

            width += image.Width;
            count++;
        }

        return width + Math.Max(0, count - 1) * HeaderFooterInlineGap;
    }

    private static double CombineHeaderFooterInlineWidths(double textWidth, double imagesWidth) {
        if (textWidth <= 0D) {
            return imagesWidth;
        }

        if (imagesWidth <= 0D) {
            return textWidth;
        }

        return textWidth + HeaderFooterInlineGap + imagesWidth;
    }

    private static double AlignHeaderFooterGroup(PdfOptions options, double groupWidth, PdfAlign align) {
        double contentLeft = options.MarginLeft;
        double contentWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        if (groupWidth > contentWidth + 0.001D) {
            throw new ArgumentException("Combined PDF header/footer content must fit inside the page content width.");
        }

        return GetHeaderFooterAlignedObjectX(contentLeft, contentWidth, groupWidth, align);
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
        bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null || shape.FillRadialGradient != null) && shape.Kind != OfficeShapeKind.Line;
        bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
        double fillOpacity = hasFill ? shape.FillOpacity ?? 1D : 1D;
        double strokeOpacity = hasStroke ? shape.StrokeOpacity ?? 1D : 1D;
        return EnsureHeaderFooterGraphicsState(page, fillOpacity, strokeOpacity);
    }

    private static string? EnsureHeaderFooterFillGradient(LayoutResult.Page page, OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
        if (shape.Kind == OfficeShapeKind.Line) return null;
        if (shape.FillRadialGradient != null) return EnsureRadialShading(page.Shadings, shape.FillRadialGradient);
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
        return EnsureAxialShading(page.Shadings, gradient, x0, y0, x1, y1);
    }

    private static void DrawHeaderFooterShapeShadowAt(StringBuilder sb, LayoutResult.Page page, OfficeShape shape, double xShape, double bottomY) {
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
                DrawHeaderFooterShapeShadowLayer(
                    sb,
                    page,
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
        DrawHeaderFooterShapeShadowLayer(sb, page, shape, shadowColor, shadowX, shadowBottomY, Math.Max(0D, shape.StrokeWidth), coreOpacity, hasFill, hasStroke);
    }

    private static void DrawHeaderFooterShapeShadowLayer(
        StringBuilder sb,
        LayoutResult.Page page,
        OfficeShape shape,
        PdfColor color,
        double x,
        double bottomY,
        double strokeWidth,
        double opacity,
        bool hasFill,
        bool hasStroke) {
        var content = new ContentStreamBuilder(sb).SaveState();
        string? graphicsState = EnsureHeaderFooterGraphicsState(page, opacity, opacity);
        if (graphicsState != null) content.GraphicsState(graphicsState);
        DrawShapeShadowLayer(sb, shape, color, x, bottomY, strokeWidth, hasFill, hasStroke);
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
            string? shadingName = EnsureHeaderFooterFillGradient(page, shape, xShape, bottomY, localCoordinates: true);
            DrawTransformedShape(sb, shape, shadingName == null ? ToHeaderFooterPdfColor(shape.FillColor) : null, ToHeaderFooterPdfColor(shape.StrokeColor), shadingName, xShape, bottomY);
        } else {
            if (shape.ClipPath != null) {
                new ContentStreamBuilder(sb)
                    .SaveState();
                AppendClipPath(sb, shape.ClipPath, xShape, bottomY, shape.Height);
            }

            string? shadingName = EnsureHeaderFooterFillGradient(page, shape, xShape, bottomY, localCoordinates: false);
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
