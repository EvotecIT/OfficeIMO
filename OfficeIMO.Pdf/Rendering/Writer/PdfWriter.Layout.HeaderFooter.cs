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
        ValidateHeaderFooterGroupLayouts(options, variantPageNumber, pageNumber, totalPages, documentPages, isHeader: true);
        ValidateHeaderFooterGroupLayouts(options, variantPageNumber, pageNumber, totalPages, documentPages, isHeader: false);
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
            System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> shapes = isHeader
                ? options.GetHeaderShapesForPage(variantPageNumber)
                : options.GetFooterShapesForPage(variantPageNumber);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, image.Align);
            double groupWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
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
        string source = isHeader ? "Header" : "Footer";
        double y = isHeader
            ? options.PageHeight - options.MarginTop + options.HeaderOffsetY - image.Height
            : options.MarginBottom - options.FooterOffsetY;
        ImageBlock block = image.ToImageBlock();
        PdfImageStyle style = block.Style ?? new PdfImageStyle();
        PageImage pageImage = CreatePageImage(block, style, x, y);
        GetImageAnnotationBounds(style, pageImage, x, y, image.Width, image.Height, out double visibleX1, out double visibleY1, out double visibleX2, out double visibleY2);
        ReportHeaderFooterBounds(
            options,
            source,
            "image",
            visibleX1,
            visibleY1,
            visibleX2 - visibleX1,
            visibleY2 - visibleY1);
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

    private static double MeasureHeaderFooterShapesWidth(System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> shapes, PdfAlign align) {
        double width = 0D;
        int count = 0;
        foreach (PdfHeaderFooterShape shape in shapes) {
            if (shape.Align != align) {
                continue;
            }

            width += shape.Width;
            count++;
        }

        return width + Math.Max(0, count - 1) * HeaderFooterInlineGap;
    }

    private static double CombineHeaderFooterInlineWidths(params double[] widths) {
        double width = 0D;
        int groups = 0;
        for (int i = 0; i < widths.Length; i++) {
            if (widths[i] <= 0D) {
                continue;
            }

            width += widths[i];
            groups++;
        }

        return width + Math.Max(0, groups - 1) * HeaderFooterInlineGap;
    }

    private static double AlignHeaderFooterGroup(PdfOptions options, double groupWidth, PdfAlign align) {
        double contentLeft = options.MarginLeft;
        double contentWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        return GetHeaderFooterAlignedObjectX(contentLeft, contentWidth, groupWidth, align);
    }

    private static string BuildHeaderFooterShapes(
        LayoutResult.Page page,
        PdfOptions options,
        int variantPageNumber,
        int pageNumber,
        int totalPages,
        int documentPages) {
        var sb = new StringBuilder();
        AddHeaderFooterShapes(
            sb,
            page,
            options,
            options.GetHeaderShapesForPage(variantPageNumber),
            variantPageNumber,
            pageNumber,
            totalPages,
            documentPages,
            isHeader: true);
        AddHeaderFooterShapes(
            sb,
            page,
            options,
            options.GetFooterShapesForPage(variantPageNumber),
            variantPageNumber,
            pageNumber,
            totalPages,
            documentPages,
            isHeader: false);

        return sb.ToString();
    }

    private static void AddHeaderFooterShapes(
        StringBuilder sb,
        LayoutResult.Page page,
        PdfOptions options,
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> shapes,
        int variantPageNumber,
        int pageNumber,
        int totalPages,
        int documentPages,
        bool isHeader) {
        var consumedWidths = new System.Collections.Generic.Dictionary<PdfAlign, double>();
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> images = isHeader
            ? options.GetHeaderImagesForPage(variantPageNumber)
            : options.GetFooterImagesForPage(variantPageNumber);
        foreach (PdfHeaderFooterShape shape in shapes) {
            double textWidth = MeasureHeaderFooterTextWidth(
                options,
                variantPageNumber,
                pageNumber,
                totalPages,
                documentPages,
                shape.Align,
                isHeader);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, shape.Align);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, shape.Align);
            double groupWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            double groupX = AlignHeaderFooterGroup(options, groupWidth, shape.Align);
            double consumedWidth = consumedWidths.TryGetValue(shape.Align, out double value) ? value : 0D;
            double shapeX = groupX +
                (textWidth > 0D ? textWidth + HeaderFooterInlineGap : 0D) +
                (imagesWidth > 0D ? imagesWidth + HeaderFooterInlineGap : 0D) +
                consumedWidth;

            AddHeaderFooterShape(sb, page, options, shape, shapeX, isHeader);
            consumedWidths[shape.Align] = consumedWidth + shape.Width + HeaderFooterInlineGap;
        }
    }

    private static void AddHeaderFooterShape(
        StringBuilder sb,
        LayoutResult.Page page,
        PdfOptions options,
        PdfHeaderFooterShape headerFooterShape,
        double x,
        bool isHeader) {
        ShapeBlock block = headerFooterShape.ToShapeBlock();
        PdfDrawingStyle style = block.Style ?? new PdfDrawingStyle();
        PdfDocument.ValidateDrawingStyle(style, isHeader ? "Header shape" : "Footer shape");

        string source = isHeader ? "Header" : "Footer";
        double bottomY = isHeader
            ? options.PageHeight - options.MarginTop + options.HeaderOffsetY - block.Shape.Height
            : options.MarginBottom - options.FooterOffsetY;
        if (TryGetHeaderFooterShapeBounds(block.Shape, x, bottomY, out double visibleX, out double visibleY, out double visibleWidth, out double visibleHeight)) {
            ReportHeaderFooterBounds(options, source, "shape", visibleX, visibleY, visibleWidth, visibleHeight);
        }

        DrawHeaderFooterShapeGeometryAt(sb, page, block.Shape, x, bottomY);
    }

    private static bool TryGetHeaderFooterShapeBounds(
        OfficeShape shape,
        double x,
        double bottomY,
        out double boundsX,
        out double boundsY,
        out double boundsWidth,
        out double boundsHeight) {
        boundsX = boundsY = boundsWidth = boundsHeight = 0D;
        if (shape.ClipPath?.Kind == OfficeClipPathKind.Empty) {
            return false;
        }

        ResolveHeaderFooterBaseGeometry(shape, out bool baseHasFill, out bool baseHasStroke);
        bool hasBounds = false;
        if (baseHasFill || baseHasStroke) {
            GetHeaderFooterShapeLayerBounds(
                shape,
                x,
                bottomY,
                baseHasStroke ? shape.StrokeWidth : 0D,
                out double baseX,
                out double baseY,
                out double baseWidth,
                out double baseHeight);
            if (shape.ClipPath == null || TryIntersectHeaderFooterShapeClipBounds(
                    shape,
                    x,
                    bottomY,
                    ref baseX,
                    ref baseY,
                    ref baseWidth,
                    ref baseHeight)) {
                IncludeShapeBounds(
                    baseX,
                    baseY,
                    baseWidth,
                    baseHeight,
                    ref hasBounds,
                    ref boundsX,
                    ref boundsY,
                    ref boundsWidth,
                    ref boundsHeight);
            }
        }

        OfficeShadow? shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) {
            return hasBounds;
        }

        double coreOpacity = shadow.Opacity * shadow.Color.A / 255D;
        ResolveShadowGeometry(shape, out bool shadowHasFill, out bool shadowHasStroke);
        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            coreOpacity,
            shadow.BlurRadius,
            shape.StrokeWidth,
            shadowHasFill,
            shadowHasStroke,
            OfficeShadowLayerPlanner.CanExpand(shape));
        double shadowX = x + shadow.OffsetX;
        double shadowBottomY = bottomY - shadow.OffsetY;
        for (int index = 0; index < layers.Count; index++) {
            OfficeShadowLayer layer = layers[index];
            OfficeShape layerShape = layer.Expansion > 0D
                ? OfficeShadowLayerPlanner.CreateExpandedShape(shape, layer.Expansion)
                : shape;
            GetHeaderFooterShapeLayerBounds(
                layerShape,
                shadowX - layer.Expansion,
                shadowBottomY - layer.Expansion,
                layer.HasStroke ? layer.StrokeWidth : 0D,
                out double layerX,
                out double layerY,
                out double layerWidth,
                out double layerHeight);
            IncludeShapeBounds(
                layerX,
                layerY,
                layerWidth,
                layerHeight,
                ref hasBounds,
                ref boundsX,
                ref boundsY,
                ref boundsWidth,
                ref boundsHeight);
        }

        return hasBounds;
    }

    private static bool TryIntersectHeaderFooterShapeClipBounds(
        OfficeShape shape,
        double x,
        double bottomY,
        ref double boundsX,
        ref double boundsY,
        ref double boundsWidth,
        ref double boundsHeight) {
        OfficeClipPath clipPath = shape.ClipPath!;
        double clipX;
        double clipY;
        double clipWidth;
        double clipHeight;
        if (!shape.Transform.HasValue) {
            clipX = x;
            clipY = bottomY + shape.Height - clipPath.Height;
            clipWidth = clipPath.Width;
            clipHeight = clipPath.Height;
        } else {
            (double left, double top, double right, double bottom) = shape.Transform.Value.TransformRectangleBounds(
                0D,
                0D,
                clipPath.Width,
                clipPath.Height);
            clipX = x + left;
            clipY = bottomY + shape.Height - bottom;
            clipWidth = right - left;
            clipHeight = bottom - top;
        }

        double intersectionX = System.Math.Max(boundsX, clipX);
        double intersectionY = System.Math.Max(boundsY, clipY);
        double intersectionRight = System.Math.Min(boundsX + boundsWidth, clipX + clipWidth);
        double intersectionTop = System.Math.Min(boundsY + boundsHeight, clipY + clipHeight);
        if (intersectionRight <= intersectionX || intersectionTop <= intersectionY) {
            return false;
        }

        boundsX = intersectionX;
        boundsY = intersectionY;
        boundsWidth = intersectionRight - intersectionX;
        boundsHeight = intersectionTop - intersectionY;
        return true;
    }

    private static void ResolveHeaderFooterBaseGeometry(OfficeShape shape, out bool hasFill, out bool hasStroke) {
        hasFill = shape.Kind != OfficeShapeKind.Line
            && (shape.FillOpacity ?? 1D) > 0D
            && ((shape.FillColor.HasValue && shape.FillColor.Value.A > 0) || shape.FillGradient != null || shape.FillRadialGradient != null);
        hasStroke = shape.StrokeWidth > 0D
            && (shape.StrokeOpacity ?? 1D) > 0D
            && shape.StrokeColor.HasValue
            && shape.StrokeColor.Value.A > 0;
    }

    private static void IncludeShapeBounds(
        double x,
        double y,
        double width,
        double height,
        ref bool hasBounds,
        ref double boundsX,
        ref double boundsY,
        ref double boundsWidth,
        ref double boundsHeight) {
        if (!hasBounds) {
            boundsX = x;
            boundsY = y;
            boundsWidth = width;
            boundsHeight = height;
            hasBounds = true;
            return;
        }

        double right = System.Math.Max(boundsX + boundsWidth, x + width);
        double top = System.Math.Max(boundsY + boundsHeight, y + height);
        boundsX = System.Math.Min(boundsX, x);
        boundsY = System.Math.Min(boundsY, y);
        boundsWidth = right - boundsX;
        boundsHeight = top - boundsY;
    }

    private static void GetHeaderFooterShapeLayerBounds(
        OfficeShape shape,
        double x,
        double bottomY,
        double strokeWidth,
        out double boundsX,
        out double boundsY,
        out double boundsWidth,
        out double boundsHeight) {
        double strokeExpansionFactor = UsesPdfMiterJoinEnvelope(shape) ? 10D : 1D;
        if (!shape.Transform.HasValue) {
            double strokeExpansion = strokeWidth * 0.5D * strokeExpansionFactor;
            boundsX = x - strokeExpansion;
            boundsY = bottomY - strokeExpansion;
            boundsWidth = shape.Width + (strokeExpansion * 2D);
            boundsHeight = shape.Height + (strokeExpansion * 2D);
            return;
        }

        OfficeTransform transform = shape.Transform.Value;
        (double left, double top, double right, double bottom) = transform.TransformRectangleBounds(0D, 0D, shape.Width, shape.Height);
        double halfStroke = strokeWidth * 0.5D * strokeExpansionFactor;
        double strokeExpansionX = halfStroke * System.Math.Sqrt((transform.M11 * transform.M11) + (transform.M21 * transform.M21));
        double strokeExpansionY = halfStroke * System.Math.Sqrt((transform.M12 * transform.M12) + (transform.M22 * transform.M22));
        boundsX = x + left - strokeExpansionX;
        boundsY = bottomY + shape.Height - bottom - strokeExpansionY;
        boundsWidth = right - left + (strokeExpansionX * 2D);
        boundsHeight = bottom - top + (strokeExpansionY * 2D);
    }

    private static bool UsesPdfMiterJoinEnvelope(OfficeShape shape) =>
        (shape.Kind == OfficeShapeKind.Polygon || shape.Kind == OfficeShapeKind.Path)
        && (!shape.StrokeLineJoin.HasValue || shape.StrokeLineJoin.Value == OfficeStrokeLineJoin.Miter);

    private static void ReportHeaderFooterBounds(
        PdfOptions options,
        string source,
        string contentKind,
        double x,
        double? y,
        double width,
        double? height) {
        const double tolerance = 0.001D;
        bool outsideVerticalBounds = y.HasValue && height.HasValue &&
            (y.Value < -tolerance || y.Value + height.Value > options.PageHeight + tolerance);
        if (x < -tolerance || x + width > options.PageWidth + tolerance || outsideVerticalBounds) {
            options.AddLayoutDiagnostic(
                "HeaderFooterPageBoundsClipped",
                source,
                source + " " + contentKind + " extends beyond the physical page bounds and may be clipped by the PDF page.",
                PdfLayoutDiagnosticKind.ClippedContent,
                x: x,
                y: y,
                width: width,
                height: height);
        }
    }

    private static double GetHeaderFooterAlignedObjectX(double containerX, double containerWidth, double objectWidth, PdfAlign align) {
        if (align == PdfAlign.Center) return containerX + ((containerWidth - objectWidth) / 2);
        if (align == PdfAlign.Right) return containerX + containerWidth - objectWidth;
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
            if (existing.FillOpacity.Equals(fillOpacity) &&
                existing.StrokeOpacity.Equals(strokeOpacity) &&
                existing.BlendMode == OfficeBlendMode.Normal) {
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
        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            coreOpacity,
            shadow.BlurRadius,
            shape.StrokeWidth,
            hasFill,
            hasStroke,
            OfficeShadowLayerPlanner.CanExpand(shape));
        for (int index = 0; index < layers.Count; index++) {
            OfficeShadowLayer layer = layers[index];
            OfficeShape layerShape = layer.Expansion > 0D
                ? OfficeShadowLayerPlanner.CreateExpandedShape(shape, layer.Expansion)
                : shape;
            DrawHeaderFooterShapeShadowLayer(
                sb,
                page,
                layerShape,
                shadowColor,
                shadowX - layer.Expansion,
                shadowBottomY - layer.Expansion,
                layer.StrokeWidth,
                layer.Opacity,
                layer.HasFill,
                layer.HasStroke);
        }
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
        if (shape.ClipPath?.Kind == OfficeClipPathKind.Empty) return;

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
