using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>
    /// Projects supported page drawing operators, text spans, and image placements into a dependency-free drawing scene.
    /// </summary>
    public OfficeDrawing ToDrawing() {
        _demandContentExtraction?.Invoke("visual content");
        (double Width, double Height) size = GetVisualPageSize();
        Matrix2D pageTransform = GetVisualPageTransform();
        var drawing = new OfficeDrawing(size.Width, size.Height);
        var textOutputBudget = CreateTextOutputBudget();
        var pageContentBudget = new PageContentBudget(this);
        var type3GlyphBudget = new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        RegisterEmbeddedFonts(drawing, ResolveDictionary(GetInheritedValue("Resources")), new HashSet<PdfStream>(), 0);

        List<PdfPageDrawingElement> pageElements = GetOrderedPageDrawingElements(size.Width, size.Height, pageTransform, textOutputBudget, pageContentBudget, type3GlyphBudget);
        IReadOnlyList<PdfPageDrawingEffectTransition> effects = GetGraphicsEffectTransitions(pageTransform, size.Height, pageContentBudget);
        var softMasks = new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
        var activeSoftMasks = new HashSet<PdfStream>();
        for (int i = 0; i < pageElements.Count; i++) {
            PdfPageDrawingElement element = pageElements[i].WithEffect(
                pageElements[i].Effect.OverlayOn(ResolveDrawingEffect(
                    effects,
                    pageElements[i].PaintOrder,
                    contentOrderKey: pageElements[i].ContentOrderKey)));
            AddDrawingElement(drawing, size.Height, pageTransform, element, softMasks, activeSoftMasks, textOutputBudget, pageContentBudget, type3GlyphBudget);
        }

        AddAnnotationAppearances(drawing, size.Height, pageTransform, textOutputBudget, pageContentBudget, type3GlyphBudget);

        return drawing;
    }

    private void RegisterEmbeddedFonts(OfficeDrawing drawing, PdfDictionary? resources, HashSet<PdfStream> activeForms, int depth) {
        EnsureContentNestingBudget(depth);
        if (resources == null) return;

        foreach (PdfFontResource font in ResourceResolver.GetFontsForResources(resources, _objects).Values) {
            if (font.EmbeddedTrueTypeFont == null) continue;
            OfficeFontInfo info = ToOfficeFontInfo(font.BaseFont, 12D, font.DrawingFontFamily);
            drawing.Fonts.TryAdd(info.FamilyName, font.EmbeddedTrueTypeFont, info.Style);
        }

        PdfDictionary? xObjects = ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) ? xObjectValue : null);
        if (xObjects == null) return;
        foreach (PdfObject value in xObjects.Items.Values) {
            if (ResolveObject(value) is not PdfStream form ||
                !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) ||
                !activeForms.Add(form)) continue;
            try {
                PdfDictionary? formResources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceValue) ? formResourceValue : null) ?? resources;
                RegisterEmbeddedFonts(drawing, formResources, activeForms, depth + 1);
            } finally {
                activeForms.Remove(form);
            }
        }
    }

    private List<PdfPageDrawingElement> GetOrderedPageDrawingElements(
        double pageWidth,
        double pageHeight,
        Matrix2D pageTransform,
        TextContentParser.TextOutputBudget textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        var elements = new List<PdfPageDrawingElement>();
        var renderedType3PaintOrders = new HashSet<double>();
        IReadOnlyList<PdfPageVisualPrimitive> primitives = GetVisualPrimitives(
            pageWidth,
            pageHeight,
            pageTransform,
            textOutputBudget,
            pageContentBudget,
            renderedType3PaintOrders,
            type3GlyphBudget,
            (placement, image, effect) => elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect)),
            (primitive, effect) => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect)),
            (group, transform, paintOrder, contentOrderKey, effect) => elements.Add(
                PdfPageDrawingElement.FromGroup(group, transform, paintOrder, contentOrderKey, elements.Count).WithEffect(effect)));
        for (int i = 0; i < primitives.Count; i++) {
            elements.Add(PdfPageDrawingElement.FromPrimitive(primitives[i], elements.Count));
        }

        IReadOnlyList<PdfTextSpan> spans = GetVisualTextSpans(pageHeight, pageTransform, textOutputBudget, pageContentBudget);
        for (int i = 0; i < spans.Count; i++) {
            if (renderedType3PaintOrders.Contains(spans[i].PaintOrder)) continue;
            elements.Add(PdfPageDrawingElement.FromText(spans[i], elements.Count));
        }

        IReadOnlyList<PdfImagePlacement> placements = GetVisualImagePlacements(pageHeight, pageTransform, pageContentBudget);
        if (placements.Count > 0) {
            IReadOnlyList<PdfExtractedImage> images = GetImages(0, placements, colorizeImageMasks: true);
            for (int i = 0; i < placements.Count; i++) {
                PdfImagePlacement placement = placements[i];
                PdfExtractedImage? image = FindImage(images, placement);
                if (image != null) {
                    elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count));
                }
            }
        }

        SortDrawingElements(elements);
        return elements;
    }

    private static void SortDrawingElements(List<PdfPageDrawingElement> elements) {
        elements.Sort(static (left, right) => {
            if (left.ContentOrderKey != null && right.ContentOrderKey != null) {
                int contentOrder = left.ContentOrderKey.CompareTo(right.ContentOrderKey);
                if (contentOrder != 0) return contentOrder;
            }
            int order = left.PaintOrder.CompareTo(right.PaintOrder);
            return order != 0 ? order : left.Sequence.CompareTo(right.Sequence);
        });
    }

    private void AddDrawingElement(
        OfficeDrawing drawing,
        double pageHeight,
        Matrix2D pageTransform,
        PdfPageDrawingElement element,
        Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask> softMasks,
        HashSet<PdfStream> activeSoftMasks,
        TextContentParser.TextOutputBudget textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        if (element.Effect.IsDefault) {
            AddDrawingElementCore(drawing, pageHeight, element);
            return;
        }

        if (element.Kind == PdfPageDrawingElementKind.Group &&
            element.GroupDrawing != null &&
            element.GroupTransform.TryInvert(out OfficeTransform inverseGroupTransform)) {
            OfficeDrawingSoftMask? groupSoftMask = null;
            if (element.Effect.SoftMask != null) {
                Matrix2D softMaskTransform = element.Effect.SoftMaskTransform ?? pageTransform;
                var inverseGroupMatrix = new Matrix2D(
                    inverseGroupTransform.M11,
                    inverseGroupTransform.M12,
                    inverseGroupTransform.M21,
                    inverseGroupTransform.M22,
                    inverseGroupTransform.OffsetX,
                    inverseGroupTransform.OffsetY);
                var parentDrawingFlip = new Matrix2D(1D, 0D, 0D, -1D, 0D, drawing.Height);
                var groupDrawingFlip = new Matrix2D(1D, 0D, 0D, -1D, 0D, element.GroupDrawing.Height);
                Matrix2D localSoftMaskTransform = Matrix2D.Multiply(
                    groupDrawingFlip,
                    Matrix2D.Multiply(
                        inverseGroupMatrix,
                        Matrix2D.Multiply(parentDrawingFlip, softMaskTransform)));
                groupSoftMask = GetOrCreateSoftMask(
                    element.Effect.SoftMask,
                    element.GroupDrawing.Width,
                    element.GroupDrawing.Height,
                    localSoftMaskTransform,
                    softMasks,
                    activeSoftMasks,
                    textOutputBudget,
                    pageContentBudget,
                    type3GlyphBudget);
            }
            drawing.AddEffectDrawing(
                element.GroupDrawing,
                element.GroupTransform,
                element.Effect.BlendMode,
                groupSoftMask);
            return;
        }

        var isolated = new OfficeDrawing(drawing.Width, drawing.Height);
        AddDrawingElementCore(isolated, pageHeight, element);
        if (isolated.Elements.Count == 0) return;
        OfficeDrawingSoftMask? softMask = element.Effect.SoftMask == null
            ? null
            : GetOrCreateSoftMask(
                element.Effect.SoftMask,
                drawing.Width,
                drawing.Height,
                element.Effect.SoftMaskTransform ?? pageTransform,
                softMasks,
                activeSoftMasks,
                textOutputBudget,
                pageContentBudget,
                type3GlyphBudget);
        drawing.AddEffectDrawing(isolated, OfficeTransform.Identity, element.Effect.BlendMode, softMask);
    }

    private static void AddDrawingElementCore(OfficeDrawing drawing, double pageHeight, PdfPageDrawingElement element) {
        switch (element.Kind) {
            case PdfPageDrawingElementKind.Primitive:
                AddVisualPrimitive(drawing, element.Primitive);
                break;
            case PdfPageDrawingElementKind.Text:
                AddTextSpan(drawing, pageHeight, element.TextSpan!);
                break;
            case PdfPageDrawingElementKind.Image:
                AddImagePlacement(drawing, pageHeight, element.ImagePlacement!, element.Image!);
                break;
            case PdfPageDrawingElementKind.Group:
                drawing.AddEffectDrawing(element.GroupDrawing!, element.GroupTransform);
                break;
        }
    }

    private static void AddVisualPrimitive(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        if (primitive.FillTilingPattern != null) {
            AddTilingPatternFill(drawing, primitive);
        }

        bool hasOrdinaryFill = primitive.FillColor.HasValue || primitive.FillGradient != null || primitive.FillRadialGradient != null;
        bool hasOrdinaryStroke = primitive.StrokeWidth > 0D &&
            (primitive.StrokeColor.HasValue || primitive.StrokeGradient != null || primitive.StrokeRadialGradient != null);
        if (hasOrdinaryFill || hasOrdinaryStroke) {
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                AddRectangle(drawing, primitive);
            } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                AddLine(drawing, primitive);
            } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Path) {
                AddPath(drawing, primitive);
            }
        }

        if (primitive.StrokeTilingPattern != null && primitive.StrokeWidth > 0D) {
            AddTilingPatternStroke(drawing, primitive);
        }
    }

    private void AddVisualPrimitives(OfficeDrawing drawing, double pageWidth, double pageHeight, Matrix2D pageTransform) {
        IReadOnlyList<PdfPageVisualPrimitive> primitives = GetVisualPrimitives(pageWidth, pageHeight, pageTransform);
        for (int i = 0; i < primitives.Count; i++) {
            PdfPageVisualPrimitive primitive = primitives[i];
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                AddRectangle(drawing, primitive);
            } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                AddLine(drawing, primitive);
            } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Path) {
                AddPath(drawing, primitive);
            }
        }
    }

    private static void AddRectangle(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        if (!HasVisibleOverlap(primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height)) {
            return;
        }

        OfficeShape shape = OfficeShape.Rectangle(primitive.Width, primitive.Height);
        shape.FillColor = primitive.FillColor;
        shape.FillGradient = primitive.FillGradient;
        shape.FillRadialGradient = primitive.FillRadialGradient;
        shape.StrokeColor = primitive.StrokeColor;
        shape.StrokeGradient = primitive.StrokeGradient;
        shape.StrokeRadialGradient = primitive.StrokeRadialGradient;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        shape.FillOpacity = primitive.FillOpacity;
        shape.StrokeOpacity = primitive.StrokeOpacity;
        shape.FillRule = primitive.FillRule;
        PdfPageClipPath? clipPath = GetEffectivePageClip(primitive.ClipPath, primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height);
        if (TryAddClippedShape(drawing, shape, primitive.X, primitive.Y, clipPath)) {
            return;
        }

        if (HasPositiveArea(primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height)) {
            drawing.AddShape(shape, primitive.X, primitive.Y);
        }
    }

    private static void AddLine(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        double left = Math.Min(primitive.X1, primitive.X2);
        double top = Math.Min(primitive.Y1, primitive.Y2);
        double right = Math.Max(primitive.X1, primitive.X2);
        double bottom = Math.Max(primitive.Y1, primitive.Y2);
        double strokeHalf = Math.Max(primitive.StrokeWidth, 1D) / 2D;
        if ((NearlyEqual(left, right) && NearlyEqual(top, bottom)) ||
            !HasVisibleOverlap(left - strokeHalf, top - strokeHalf, right - left + (strokeHalf * 2D), bottom - top + (strokeHalf * 2D), drawing.Width, drawing.Height)) {
            return;
        }

        OfficeShape shape = OfficeShape.Line(primitive.X1 - left, primitive.Y1 - top, primitive.X2 - left, primitive.Y2 - top);
        shape.StrokeColor = primitive.StrokeColor;
        shape.StrokeGradient = primitive.StrokeGradient;
        shape.StrokeRadialGradient = primitive.StrokeRadialGradient;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        shape.StrokeOpacity = primitive.StrokeOpacity;
        PdfPageClipPath? clipPath = GetEffectivePageClip(primitive.ClipPath, left - strokeHalf, top - strokeHalf, right - left + (strokeHalf * 2D), bottom - top + (strokeHalf * 2D), drawing.Width, drawing.Height);
        if (TryAddClippedShape(drawing, shape, left, top, clipPath)) {
            return;
        }

        if (HasVisibleOverlap(left - strokeHalf, top - strokeHalf, right - left + (strokeHalf * 2D), bottom - top + (strokeHalf * 2D), drawing.Width, drawing.Height)) {
            drawing.AddShape(shape, left, top);
        }
    }

    private static void AddPath(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        if (!HasVisibleOverlap(primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height)) {
            return;
        }

        OfficeShape shape = OfficeShape.Path(primitive.PathCommands);
        shape.FillColor = primitive.FillColor;
        shape.FillGradient = primitive.FillGradient;
        shape.FillRadialGradient = primitive.FillRadialGradient;
        shape.StrokeColor = primitive.StrokeColor;
        shape.StrokeGradient = primitive.StrokeGradient;
        shape.StrokeRadialGradient = primitive.StrokeRadialGradient;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        shape.FillOpacity = primitive.FillOpacity;
        shape.StrokeOpacity = primitive.StrokeOpacity;
        shape.FillRule = primitive.FillRule;
        PdfPageClipPath? clipPath = GetEffectivePageClip(primitive.ClipPath, primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height);
        if (TryAddClippedShape(drawing, shape, primitive.X, primitive.Y, clipPath)) {
            return;
        }

        if (HasPositiveArea(primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height)) {
            drawing.AddShape(shape, primitive.X, primitive.Y);
        }
    }

    private static PdfPageClipPath? GetEffectivePageClip(PdfPageClipPath? clipPath, double x, double y, double width, double height, double drawingWidth, double drawingHeight) {
        PdfPageClipPath pageClip = PdfPageClipPath.Rectangle(0D, 0D, drawingWidth, drawingHeight);
        if (!clipPath.HasValue) {
            return HasPositiveArea(x, y, width, height, drawingWidth, drawingHeight) ? null : pageClip;
        }

        if (!clipPath.Value.IsRectangle) {
            return clipPath;
        }

        return IntersectClipBounds(clipPath.Value, pageClip, out PdfPageClipPath intersection)
            ? intersection
            : PdfPageClipPath.Rectangle(0D, 0D, 0D, 0D);
    }

    private static bool IntersectClipBounds(PdfPageClipPath first, PdfPageClipPath second, out PdfPageClipPath intersection) {
        double left = Math.Max(first.X, second.X);
        double top = Math.Max(first.Y, second.Y);
        double right = Math.Min(first.X + first.Width, second.X + second.Width);
        double bottom = Math.Min(first.Y + first.Height, second.Y + second.Height);
        double width = right - left;
        double height = bottom - top;
        if (width <= 0D || height <= 0D) {
            intersection = default;
            return false;
        }

        intersection = PdfPageClipPath.Rectangle(left, top, width, height);
        return true;
    }

    private static bool TryAddClippedShape(OfficeDrawing drawing, OfficeShape shape, double x, double y, PdfPageClipPath? clipPath) {
        if (!clipPath.HasValue) {
            return false;
        }

        PdfPageClipPath clip = clipPath.Value;
        if (clip.Width <= 0D || clip.Height <= 0D) {
            return true;
        }

        if (!TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out PdfPageClipPath drawingClip)) {
            return true;
        }

        clip = drawingClip;
        OfficeClipPath? localClip = clip.ToOfficeClipPath(x, y);
        if (localClip != null && HasPositiveArea(x, y, shape.Width, shape.Height, drawing.Width, drawing.Height)) {
            shape.ClipPath = localClip;
            return false;
        }

        OfficeClipPath? groupClip = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (groupClip == null) {
            return false;
        }

        double localX = x - clip.X;
        double localY = y - clip.Y;
        double shapeX = localX;
        double shapeY = localY;
        if (shapeX < 0D || shapeY < 0D) {
            double translatedX = Math.Max(0D, shapeX);
            double translatedY = Math.Max(0D, shapeY);
            double offsetX = shapeX - translatedX;
            double offsetY = shapeY - translatedY;
            shape = shape.Clone();
            OfficeTransform offsetTransform = OfficeTransform.Translate(offsetX, offsetY);
            shape.Transform = shape.Transform.HasValue ? offsetTransform.Then(shape.Transform.Value) : offsetTransform;
            shapeX = translatedX;
            shapeY = translatedY;
        }

        double innerWidth = Math.Max(clip.Width, shapeX + shape.Width);
        double innerHeight = Math.Max(clip.Height, shapeY + shape.Height);
        var innerDrawing = new OfficeDrawing(innerWidth, innerHeight);
        innerDrawing.AddShape(shape, shapeX, shapeY);
        drawing.AddClippedDrawing(innerDrawing, clip.X, clip.Y, groupClip);
        return true;
    }

    private IReadOnlyList<PdfPageVisualPrimitive> GetVisualPrimitives(
        double pageWidth,
        double pageHeight,
        Matrix2D pageTransform,
        TextContentParser.TextOutputBudget? textOutputBudget = null,
        PageContentBudget? pageContentBudget = null,
        HashSet<double>? renderedType3PaintOrders = null,
        Type3GlyphBudget? type3GlyphBudget = null,
        Action<PdfImagePlacement, PdfExtractedImage, PdfPageDrawingEffect>? type3ImageVisitor = null,
        Action<PdfPageVisualPrimitive, PdfPageDrawingEffect>? type3PrimitiveVisitor = null,
        Action<OfficeDrawing, OfficeTransform, double, PdfContentOrderKey?, PdfPageDrawingEffect>? type3GroupVisitor = null) {
        textOutputBudget ??= CreateTextOutputBudget();
        pageContentBudget ??= new PageContentBudget(this);
        var primitives = new List<PdfPageVisualPrimitive>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        var activeType3Glyphs = new HashSet<PdfStream>();
        renderedType3PaintOrders ??= new HashSet<double>();
        type3GlyphBudget ??= new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        var tilingPatternResourceCache = new TilingPatternResourceCache();
        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            CollectVisualPrimitivesAndForms(
                content,
                pageResources,
                pageTransform,
                pageWidth,
                pageHeight,
                primitives.Add,
                activeForms,
                activeType3Glyphs: activeType3Glyphs,
                renderedType3PaintOrders: renderedType3PaintOrders,
                type3GlyphBudget: type3GlyphBudget,
                type3ImageVisitor: type3ImageVisitor,
                type3PrimitiveVisitor: type3PrimitiveVisitor,
                type3GroupVisitor: type3GroupVisitor,
                tilingPatternResourceCache: tilingPatternResourceCache,
                textOutputBudget: textOutputBudget,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root);
        }

        return primitives.Count == 0 ? Array.Empty<PdfPageVisualPrimitive>() : primitives.AsReadOnly();
    }

    private void CollectVisualPrimitivesAndForms(
        string content,
        PdfDictionary? resources,
        Matrix2D baseTransform,
        double pageWidth,
        double pageHeight,
        Action<PdfPageVisualPrimitive> primitiveVisitor,
        HashSet<PdfStream> activeForms,
        HashSet<PdfStream>? activeType3Glyphs = null,
        HashSet<double>? renderedType3PaintOrders = null,
        Type3GlyphBudget? type3GlyphBudget = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        PdfPagePatternSelection? initialFillPattern = null,
        PdfPageColorSpace? initialFillPatternBaseColorSpace = null,
        double? initialFillOpacity = null,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        PdfPagePatternSelection? initialStrokePattern = null,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace = null,
        double? initialStrokeOpacity = null,
        double? initialStrokeWidth = null,
        OfficeStrokeDashStyle? initialStrokeDashStyle = null,
        OfficeStrokeLineCap? initialStrokeLineCap = null,
        OfficeStrokeLineJoin? initialStrokeLineJoin = null,
        int contentNestingDepth = 0,
        bool includeTilingPatterns = true,
        bool retainPrimitiveData = true,
        bool requireSupportedType3Content = false,
        bool allowSupportedType3Patterns = false,
        bool allowSupportedType3TransparencyGroups = false,
        bool requireNestedType3Uncolored = false,
        Action<string>? unrenderedPatternVisitor = null,
        Action<PdfImagePlacement, PdfExtractedImage, PdfPageDrawingEffect>? type3ImageVisitor = null,
        Action<PdfPageVisualPrimitive, PdfPageDrawingEffect>? type3PrimitiveVisitor = null,
        Action<OfficeDrawing, OfficeTransform, double, PdfContentOrderKey?, PdfPageDrawingEffect>? type3GroupVisitor = null,
        Action<PdfPageGraphicsStateResource, Matrix2D>? graphicsStateVisitor = null,
        TilingPatternResourceCache? tilingPatternResourceCache = null,
        TextContentParser.TextOutputBudget? textOutputBudget = null,
        PageContentBudget? pageContentBudget = null,
        PdfContentOrderKey? contentOrderPrefix = null) {
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        activeType3Glyphs ??= new HashSet<PdfStream>();
        renderedType3PaintOrders ??= new HashSet<double>();
        type3GlyphBudget ??= new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        Dictionary<string, Func<byte[], double>> widthProviders = resources == null
            ? new Dictionary<string, Func<byte[], double>>(StringComparer.Ordinal)
            : ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
        Dictionary<string, PdfPageColorSpace> colorSpaceResources = GetColorSpaceResources(resources);
        Dictionary<string, PdfPageColorSpace> patternBaseColorSpaces = GetPatternBaseColorSpaceResources(resources);
        var invokedPatternNames = new HashSet<string>(StringComparer.Ordinal);
        var type3PaintChannelCache = new Type3PaintChannelCache();
        var activeType3PaintChannelStreams = new HashSet<PdfStream>();
        if (includeTilingPatterns) {
            _ = PdfPageXObjectInvocationParser.Parse(
                content,
                baseTransform,
                pageHeight,
                GetGraphicsStateResources(resources),
                colorSpaceResources,
                GetOptionalContentVisibility(resources),
                paintOrderBase: paintOrderBase,
                paintOrderScale: paintOrderScale,
                paintOrderOffset: paintOrderOffset,
                initialClipPath: initialClipPath,
                initialFillColor: initialFillColor,
                initialFillColorSpace: initialFillColorSpace,
                initialFillOpacity: initialFillOpacity,
                initialStrokeColor: initialStrokeColor,
                initialStrokeColorSpace: initialStrokeColorSpace,
                initialStrokeOpacity: initialStrokeOpacity,
                initialStrokeWidth: initialStrokeWidth,
                initialStrokeDashStyle: initialStrokeDashStyle,
                initialStrokeLineCap: initialStrokeLineCap,
                initialStrokeLineJoin: initialStrokeLineJoin,
                maxOperations: _limits.MaxContentOperations,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                fonts: fonts,
                fontWidthProviders: widthProviders,
                patternInvocationVisitor: name => invokedPatternNames.Add(name),
                patternBaseColorSpaces: patternBaseColorSpaces,
                initialFillPattern: initialFillPattern,
                initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
                initialStrokePattern: initialStrokePattern,
                initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
                type3PaintChannelResolver: glyph => ResolveType3PaintChannels(
                    glyph,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget,
                    type3GlyphBudget),
                xObjectPaintChannelResolver: (name, paintState) => ResolveXObjectPaintChannels(
                    resources,
                    name,
                    paintState,
                    pageWidth,
                    pageHeight,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget,
                    type3GlyphBudget),
                pageWidth: pageWidth);
        }
        Dictionary<string, PdfPageShadingPatternResource> shadingPatternResources = GetShadingPatternResources(resources);
        Dictionary<string, PdfPageShadingResource> shadingResources = GetShadingResources(resources);
        Dictionary<string, PdfPageTilingPatternResource>? tilingPatternResources = includeTilingPatterns
            ? GetTilingPatternResources(
                resources,
                invokedPatternNames,
                tilingPatternResourceCache,
                textOutputBudget,
                pageContentBudget,
                type3GlyphBudget,
                requireSupportedType3Content,
                contentNestingDepth,
                allowNestedPatternContent: allowSupportedType3Patterns)
            : null;
        string transformedContent = WrapContentWithTransform(content, baseTransform, out int transformedContentOffset);
        Action<PdfPageVisualPrimitive> currentPrimitiveVisitor = contentOrderPrefix == null
            ? primitiveVisitor
            : primitive => primitiveVisitor(primitive.WithContentOrderKey(contentOrderPrefix.Append(primitive.SourceOperatorIndex - transformedContentOffset)));
        _ = PdfPageContentVisualParser.Parse(
            transformedContent,
            pageWidth,
            pageHeight,
            GetGraphicsStateResources(resources),
            colorSpaceResources,
            shadingResources,
            shadingPatternResources,
            tilingPatternResources,
            GetOptionalContentVisibility(resources),
            paintOrderBase,
            paintOrderScale,
            paintOrderOffset - transformedContentOffset,
            initialClipPath,
            initialFillColor,
            initialFillColorSpace,
            initialFillOpacity,
            initialStrokeColor,
            initialStrokeColorSpace,
            initialStrokeOpacity,
            initialStrokeWidth,
            initialStrokeDashStyle,
            initialStrokeLineCap,
            initialStrokeLineJoin,
            maxOperations: _limits.MaxContentOperations,
            patternBaseColorSpaces: patternBaseColorSpaces,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            primitiveVisitor: currentPrimitiveVisitor,
            retainPrimitiveData: retainPrimitiveData,
            scaleStrokeWidthWithTransform: requireSupportedType3Content,
            unsupportedShadingTransformVisitor: requireSupportedType3Content
                ? type3GlyphBudget.RecordFailure
                : null,
            requireExactType3ShadingProjection: requireSupportedType3Content,
            authoredShadingInvocationVisitor: requireSupportedType3Content
                ? name => {
                    if (requireNestedType3Uncolored ||
                        !shadingResources.TryGetValue(name, out PdfPageShadingResource shading) ||
                        !shading.SupportsExactType3Projection) {
                        type3GlyphBudget.RecordFailure();
                    }
                }
                : null,
            unrenderedShadingVisitor: requireSupportedType3Content
                ? _ => type3GlyphBudget.RecordFailure()
                : null,
            unsupportedOperatorVisitor: requireSupportedType3Content
                ? _ => type3GlyphBudget.RecordFailure()
                : null,
            initialFillPattern: initialFillPattern,
            initialStrokePattern: initialStrokePattern);

        foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                     content,
                     baseTransform,
                     pageHeight,
                     GetGraphicsStateResources(resources),
                     colorSpaceResources,
                      GetOptionalContentVisibility(resources),
                      paintOrderBase: paintOrderBase,
                      paintOrderScale: paintOrderScale,
                      paintOrderOffset: paintOrderOffset,
                      initialClipPath: initialClipPath,
                      initialFillColor: initialFillColor,
                      initialFillColorSpace: initialFillColorSpace,
                      initialFillOpacity: initialFillOpacity,
                      initialStrokeColor: initialStrokeColor,
                      initialStrokeColorSpace: initialStrokeColorSpace,
                      initialStrokeOpacity: initialStrokeOpacity,
                      initialStrokeWidth: initialStrokeWidth,
                      initialStrokeDashStyle: initialStrokeDashStyle,
                      initialStrokeLineCap: initialStrokeLineCap,
                      initialStrokeLineJoin: initialStrokeLineJoin,
                      maxOperations: _limits.MaxContentOperations,
                      maxNestingDepth: _limits.MaxContentNestingDepth,
                      maxOperands: _limits.MaxContentOperands,
                      fonts: fonts,
                      fontWidthProviders: widthProviders,
                      type3TextVisitor: invocation => {
                          if (requireNestedType3Uncolored) {
                              for (int glyphIndex = 0; glyphIndex < invocation.Glyphs.Count; glyphIndex++) {
                                  if (invocation.Glyphs[glyphIndex].Font.Type3 is not PdfType3FontResource nestedType3 ||
                                      !nestedType3.IsUncolored) {
                                      type3GlyphBudget.RecordFailure();
                                      return false;
                                  }
                              }
                          }
                          int glyphContentNestingDepth = activeType3Glyphs.Count == 0
                              ? contentNestingDepth + 1
                              : contentNestingDepth;
                          bool rendered = RenderType3TextInvocation(
                              invocation,
                              pageWidth,
                              pageHeight,
                              primitiveVisitor,
                              activeForms,
                              activeType3Glyphs,
                              type3GlyphBudget,
                              paintOrderScale,
                              includeTilingPatterns,
                              retainPrimitiveData,
                              tilingPatternResourceCache,
                              textOutputBudget,
                              pageContentBudget,
                              glyphContentNestingDepth,
                              type3ImageVisitor,
                              type3PrimitiveVisitor,
                              type3GroupVisitor,
                              contentOrderPrefix?.Append(invocation.SourceOperatorIndex));
                          if (!rendered) type3GlyphBudget.RecordFailure();
                          return rendered;
                      },
                      renderedType3PaintOrders: renderedType3PaintOrders,
                      type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
                      unsupportedTextVisitor: requireSupportedType3Content ? type3GlyphBudget.RecordFailure : null,
                      unsupportedGraphicsEffectVisitor: requireSupportedType3Content ? type3GlyphBudget.RecordFailure : null,
                      unsupportedColorVisitor: requireSupportedType3Content ? type3GlyphBudget.RecordFailure : null,
                      patternInvocationVisitor: requireSupportedType3Content || unrenderedPatternVisitor != null
                          ? name => {
                              unrenderedPatternVisitor?.Invoke(name);
                              if ((initialFillPattern.HasValue && string.Equals(initialFillPattern.Value.Name, name, StringComparison.Ordinal)) ||
                                  (initialStrokePattern.HasValue && string.Equals(initialStrokePattern.Value.Name, name, StringComparison.Ordinal))) {
                                  return;
                              }
                              if (!allowSupportedType3Patterns ||
                                  !(shadingPatternResources.ContainsKey(name) || tilingPatternResources?.ContainsKey(name) == true)) {
                                  if (requireSupportedType3Content) type3GlyphBudget.RecordFailure();
                              }
                          }
                          : null,
                      authoredPatternInvocationVisitor: requireNestedType3Uncolored
                          ? _ => type3GlyphBudget.RecordFailure()
                          : null,
                      graphicsStateVisitor: graphicsStateVisitor,
                      allowSupportedGraphicsEffects: requireSupportedType3Content,
                      patternBaseColorSpaces: patternBaseColorSpaces,
                      initialFillPattern: initialFillPattern,
                      initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
                      initialStrokePattern: initialStrokePattern,
                      initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
                      tilingPatterns: tilingPatternResources,
                      shadingPatterns: shadingPatternResources,
                      type3PaintChannelResolver: glyph => ResolveType3PaintChannels(
                          glyph,
                          type3PaintChannelCache,
                          activeType3PaintChannelStreams,
                          pageContentBudget,
                          type3GlyphBudget),
                      xObjectPaintChannelResolver: (name, paintState) => ResolveXObjectPaintChannels(
                          resources,
                          name,
                          paintState,
                          pageWidth,
                          pageHeight,
                          type3PaintChannelCache,
                          activeType3PaintChannelStreams,
                          pageContentBudget,
                          type3GlyphBudget),
                      softMaskVisibilityResolver: (softMask, transform) => !IsSoftMaskEntirelyTransparent(
                          softMask,
                          transform,
                          resources,
                          pageWidth,
                          pageHeight,
                          type3PaintChannelCache,
                          activeType3PaintChannelStreams,
                          pageContentBudget,
                          type3GlyphBudget,
                          contentNestingDepth + 1),
                      invalidPatternSelectionVisitor: requireSupportedType3Content
                          ? type3GlyphBudget.RecordFailure
                          : null,
                      pageWidth: pageWidth)) {
            if (!TryGetFormStream(resources, invocation.Name, out PdfStream formStream)) {
                if (requireSupportedType3Content && invocation.InlineImage == null && !TryGetImageXObject(resources, invocation.Name, out _, out _)) {
                    type3GlyphBudget.RecordFailure();
                }
                continue;
            }

            if (requireSupportedType3Content &&
                ResolveXObjectPaintChannels(
                    resources,
                    invocation.Name,
                    invocation.PaintState,
                    pageWidth,
                    pageHeight,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget,
                    type3GlyphBudget) == PdfType3PaintChannels.None) {
                continue;
            }

            if (requireSupportedType3Content && formStream.Dictionary.Items.ContainsKey("OC")) {
                type3GlyphBudget.RecordFailure();
                continue;
            }

            if (!activeForms.Add(formStream)) {
                if (requireSupportedType3Content) type3GlyphBudget.RecordFailure();
                continue;
            }

            try {
                PdfDictionary formDictionary = formStream.Dictionary;
                PdfDictionary? formResources = ResolveDictionary(formDictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? resources;
                Matrix2D formTransform = ApplyFormMatrix(invocation.Transform, formDictionary);
                PdfContentOrderKey? formOrderPrefix = contentOrderPrefix?.Append(invocation.SourceOperatorIndex);
                bool projectsType3TransparencyGroup = requireSupportedType3Content && formDictionary.Items.ContainsKey("Group");
                if (projectsType3TransparencyGroup) {
                    if ((invocation.FillOpacity ?? 1D) <= 0D) {
                        continue;
                    }
                    Type3TransparencyGroupDrawingResult boundsResult = TryGetVisibleType3TransparencyGroupBounds(
                        formDictionary,
                        formTransform,
                        invocation.ClipPath,
                        pageWidth,
                        pageHeight,
                        out _);
                    if (boundsResult == Type3TransparencyGroupDrawingResult.Invisible) {
                        continue;
                    }
                    if (boundsResult == Type3TransparencyGroupDrawingResult.Unsupported ||
                        !allowSupportedType3TransparencyGroups ||
                        type3GroupVisitor == null ||
                        !IsSupportedType3TransparencyGroup(formDictionary)) {
                        type3GlyphBudget.RecordFailure();
                        continue;
                    }
                }
                string decodedFormContent = PdfEncoding.Latin1GetString(pageContentBudget.Decode(formStream));
                string formContent = WrapFormContentWithBoundingBoxClip(decodedFormContent, formDictionary);
                if (projectsType3TransparencyGroup) {
                    Type3TransparencyGroupDrawingResult groupResult = TryCreateType3TransparencyGroupDrawing(
                            decodedFormContent,
                            formDictionary,
                            formResources,
                            formTransform,
                            pageWidth,
                            pageHeight,
                            invocation,
                            activeForms,
                            activeType3Glyphs,
                            renderedType3PaintOrders,
                              type3GlyphBudget,
                              paintOrderScale,
                              includeTilingPatterns,
                              retainPrimitiveData,
                              requireNestedType3Uncolored,
                              tilingPatternResourceCache,
                            textOutputBudget,
                            pageContentBudget,
                            contentNestingDepth,
                            formOrderPrefix,
                            out OfficeDrawing? groupDrawing,
                            out OfficeTransform groupTransform);
                    if (groupResult == Type3TransparencyGroupDrawingResult.Unsupported) {
                        type3GlyphBudget.RecordFailure();
                        continue;
                    }
                    if (groupResult == Type3TransparencyGroupDrawingResult.Invisible) continue;
                    type3GroupVisitor!(groupDrawing, groupTransform, invocation.PaintOrder, formOrderPrefix, PdfPageDrawingEffect.Default);
                    continue;
                }
                CollectVisualPrimitivesAndForms(
                    formContent,
                    formResources,
                    formTransform,
                    pageWidth,
                    pageHeight,
                    primitiveVisitor,
                    activeForms,
                    activeType3Glyphs,
                    renderedType3PaintOrders,
                    type3GlyphBudget,
                    invocation.PaintOrder,
                    paintOrderScale * 0.000000001D,
                    initialClipPath: invocation.ClipPath,
                    initialFillColor: invocation.FillColor,
                    initialFillColorSpace: invocation.FillColorSpace,
                    initialFillPattern: invocation.FillPattern,
                    initialFillPatternBaseColorSpace: invocation.FillPatternBaseColorSpace,
                    initialFillOpacity: invocation.FillOpacity,
                    initialStrokeColor: invocation.StrokeColor,
                    initialStrokeColorSpace: invocation.StrokeColorSpace,
                    initialStrokePattern: invocation.StrokePattern,
                    initialStrokePatternBaseColorSpace: invocation.StrokePatternBaseColorSpace,
                    initialStrokeOpacity: invocation.StrokeOpacity,
                    initialStrokeWidth: invocation.StrokeWidth,
                    initialStrokeDashStyle: invocation.StrokeDashStyle,
                    initialStrokeLineCap: invocation.StrokeLineCap,
                    initialStrokeLineJoin: invocation.StrokeLineJoin,
                    contentNestingDepth: contentNestingDepth + 1,
                    includeTilingPatterns: includeTilingPatterns,
                    retainPrimitiveData: retainPrimitiveData,
                    requireSupportedType3Content: requireSupportedType3Content,
                    allowSupportedType3Patterns: allowSupportedType3Patterns,
                    allowSupportedType3TransparencyGroups: allowSupportedType3TransparencyGroups,
                    requireNestedType3Uncolored: requireNestedType3Uncolored,
                    unrenderedPatternVisitor: unrenderedPatternVisitor,
                    type3ImageVisitor: type3ImageVisitor,
                    type3PrimitiveVisitor: type3PrimitiveVisitor,
                    type3GroupVisitor: type3GroupVisitor,
                    graphicsStateVisitor: graphicsStateVisitor,
                    tilingPatternResourceCache: tilingPatternResourceCache,
                    textOutputBudget: textOutputBudget,
                    pageContentBudget: pageContentBudget,
                    contentOrderPrefix: formOrderPrefix);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private Dictionary<string, PdfPageShadingResource> GetShadingResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageShadingResource>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("Shading", out PdfObject? shadingObject)) {
            return result;
        }

        PdfDictionary? shadings = ResolveDictionary(shadingObject);
        if (shadings == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in shadings.Items) {
            if (TryReadShading(entry.Value, out PdfPageShadingResource shading)) {
                result[entry.Key] = shading;
            }
        }

        return result;
    }

    private Dictionary<string, PdfPageShadingPatternResource> GetShadingPatternResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageShadingPatternResource>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("Pattern", out PdfObject? patternObject)) {
            return result;
        }

        PdfDictionary? patterns = ResolveDictionary(patternObject);
        if (patterns == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in patterns.Items) {
            if (TryReadShadingPattern(entry.Value, out PdfPageShadingPatternResource pattern)) {
                result[entry.Key] = pattern;
            } else if (TryReadInteger(
                ResolveDictionary(entry.Value)?.Items.TryGetValue("PatternType", out PdfObject? patternTypeObject) == true
                    ? patternTypeObject
                    : null) == 2) {
                result[entry.Key] = PdfPageShadingPatternResource.Unsupported;
            }
        }

        return result;
    }

    private bool TryReadShadingPattern(PdfObject? value, out PdfPageShadingPatternResource pattern) {
        pattern = default;
        PdfDictionary? dictionary = ResolveDictionary(value);
        if (dictionary == null ||
            TryReadInteger(dictionary.Items.TryGetValue("PatternType", out PdfObject? patternTypeObject) ? patternTypeObject : null) != 2 ||
            HasUnsupportedShadingPatternGraphicsState(dictionary) ||
            !dictionary.Items.TryGetValue("Shading", out PdfObject? shadingObject) ||
            !TryReadShading(shadingObject, out PdfPageShadingResource shading)) {
            return false;
        }

        bool hasExactMatrix = true;
        Matrix2D matrix;
        if (dictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject)) {
            matrix = ReadPatternMatrix(matrixObject);
            hasExactMatrix = TryReadStrictPatternMatrix(matrixObject, out _);
        } else {
            matrix = Matrix2D.Identity;
        }
        pattern = new PdfPageShadingPatternResource(shading, matrix, hasExactMatrix);
        return true;
    }

    private bool HasUnsupportedShadingPatternGraphicsState(PdfDictionary pattern) {
        if (!pattern.Items.TryGetValue("ExtGState", out PdfObject? value)) return false;
        PdfObject? resolved = ResolveEffectObject(value);
        return resolved is not PdfNull && resolved is not PdfDictionary { Items.Count: 0 };
    }

    private Matrix2D ReadPatternMatrix(PdfObject? matrixObject) {
        PdfArray? matrix = ResolveArray(matrixObject);
        if (matrix == null || matrix.Items.Count < 6) {
            return Matrix2D.Identity;
        }

        return new Matrix2D(
            ReadMatrixNumber(matrix, 0, 1D),
            ReadMatrixNumber(matrix, 1, 0D),
            ReadMatrixNumber(matrix, 2, 0D),
            ReadMatrixNumber(matrix, 3, 1D),
            ReadMatrixNumber(matrix, 4, 0D),
            ReadMatrixNumber(matrix, 5, 0D));
    }

    private bool TryReadStrictPatternMatrix(PdfObject? matrixObject, out Matrix2D matrix) {
        matrix = Matrix2D.Identity;
        PdfArray? values = ResolveArray(matrixObject);
        if (values == null || values.Items.Count != 6) return false;
        var components = new double[6];
        for (int index = 0; index < components.Length; index++) {
            if (ResolveObject(values.Items[index]) is not PdfNumber number || !IsFinite(number.Value)) return false;
            components[index] = number.Value;
        }
        matrix = new Matrix2D(components[0], components[1], components[2], components[3], components[4], components[5]);
        return true;
    }

    private bool TryReadShading(PdfObject? value, out PdfPageShadingResource shading) {
        shading = default;
        PdfDictionary? dictionary = ResolveDictionary(value);
        if (dictionary == null ||
            !dictionary.Items.TryGetValue("Coords", out PdfObject? coordsObject)) {
            return false;
        }

        int? shadingType = TryReadInteger(dictionary.Items.TryGetValue("ShadingType", out PdfObject? shadingTypeObject) ? shadingTypeObject : null);
        IReadOnlyList<double> coords = ReadNumberArray(coordsObject);
        if ((shadingType == 2 && coords.Count < 4) ||
            (shadingType == 3 && coords.Count < 6)) {
            return false;
        }

        if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ||
            !TryReadColorSpaceResource(colorSpaceObject, out PdfPageColorSpace colorSpace)) return false;

        bool exactColorInterpolation =
            colorSpace.Kind is PdfPageColorSpaceKind.DeviceGray or PdfPageColorSpaceKind.DeviceRgb &&
            !colorSpace.UsesIccApproximation;
        bool hasShadingBoundingBox = dictionary.Items.TryGetValue("BBox", out PdfObject? shadingBoxObject) &&
            ResolveObject(shadingBoxObject) is not PdfNull;

        if (dictionary.Items.TryGetValue("Domain", out PdfObject? shadingDomainObject) &&
            !IsCanonicalUnitIntervals(ReadNumberArray(shadingDomainObject), 1)) return false;

        bool extendsBothEnds = false;
        if (dictionary.Items.TryGetValue("Extend", out PdfObject? extendObject) &&
            ResolveObject(extendObject) is not PdfNull) {
            PdfArray? extend = ResolveArray(extendObject);
            if (extend == null || extend.Items.Count != 2 ||
                ResolveObject(extend.Items[0]) is not PdfBoolean extendStart ||
                ResolveObject(extend.Items[1]) is not PdfBoolean extendEnd) return false;
            extendsBothEnds = extendStart.Value && extendEnd.Value;
        }

        PdfObject? functionObject = dictionary.Items.TryGetValue("Function", out PdfObject? authoredFunction) ? authoredFunction : null;
        if (!TryReadShadingStops(functionObject, colorSpace, out IReadOnlyList<OfficeGradientStop> stops)) {
            return false;
        }
        exactColorInterpolation &= HasExactType3FunctionComponentRange(functionObject, colorSpace.ComponentCount);

        if (shadingType == 2) {
            bool exactAxialFamily = coords.Count == 4 && coords.All(IsFinite);
            shading = new PdfPageShadingResource(
                coords[0],
                coords[1],
                coords[2],
                coords[3],
                stops,
                exactColorInterpolation && exactAxialFamily && !hasShadingBoundingBox && extendsBothEnds);
            return true;
        }

        if (shadingType == 3) {
            bool exactRadialFamily = coords.Take(6).All(IsFinite) &&
                coords[2] >= 0D &&
                coords[5] >= 0D &&
                coords[0].Equals(coords[3]) &&
                coords[1].Equals(coords[4]) &&
                coords[5] > coords[2];
            shading = new PdfPageShadingResource(
                coords[0],
                coords[1],
                Math.Max(0D, coords[2]),
                coords[3],
                coords[4],
                Math.Max(0D, coords[5]),
                stops,
                exactColorInterpolation && !hasShadingBoundingBox && extendsBothEnds && exactRadialFamily);
            return true;
        }

        return false;
    }

    private bool HasExactType3FunctionComponentRange(PdfObject? functionObject, int componentCount) {
        PdfObject? resolved = ResolveObject(functionObject);
        if (resolved is PdfArray functionArray) {
            if (functionArray.Items.Count != 1) return false;
            resolved = ResolveObject(functionArray.Items[0]);
        }
        if (resolved is not PdfDictionary function) return false;

        int? functionType = TryReadInteger(function.Items.TryGetValue("FunctionType", out PdfObject? typeObject) ? typeObject : null);
        if (functionType == 2) {
            IReadOnlyList<double> c0 = function.Items.TryGetValue("C0", out PdfObject? c0Object)
                ? ReadNumberArray(c0Object)
                : new[] { 0D };
            IReadOnlyList<double> c1 = function.Items.TryGetValue("C1", out PdfObject? c1Object)
                ? ReadNumberArray(c1Object)
                : new[] { 1D };
            return c0.Count == componentCount &&
                   c1.Count == componentCount &&
                   c0.All(IsUnitComponent) &&
                   c1.All(IsUnitComponent);
        }

        if (functionType != 3 ||
            !function.Items.TryGetValue("Functions", out PdfObject? functionsObject) ||
            ResolveArray(functionsObject) is not PdfArray functions ||
            functions.Items.Count < 2 ||
            functions.Items.Count > 32) return false;
        return functions.Items.All(item => HasExactType3FunctionComponentRange(item, componentCount));
    }

    private static bool IsUnitComponent(double value) => IsFinite(value) && value >= 0D && value <= 1D;

    private bool TryReadShadingStops(PdfObject? functionObject, PdfPageColorSpace colorSpace, out IReadOnlyList<OfficeGradientStop> stops) {
        stops = Array.Empty<OfficeGradientStop>();
        PdfObject? resolved = ResolveObject(functionObject);
        if (resolved is PdfArray functionArray) {
            if (functionArray.Items.Count != 1) return false;
            resolved = ResolveObject(functionArray.Items[0]);
        }
        if (resolved is not PdfDictionary function) return false;

        int? functionType = TryReadInteger(function.Items.TryGetValue("FunctionType", out PdfObject? typeObject) ? typeObject : null);
        if (functionType == 2) {
            if (!TryReadType2FunctionColors(function, colorSpace, reversed: false, out OfficeColor start, out OfficeColor end)) return false;
            stops = new[] { new OfficeGradientStop(0D, start), new OfficeGradientStop(1D, end) };
            return true;
        }
        if (functionType != 3 || !function.Items.TryGetValue("Functions", out PdfObject? functionsObject)) return false;

        PdfArray? functions = ResolveArray(functionsObject);
        if (functions == null || functions.Items.Count < 2 || functions.Items.Count > 32) return false;
        IReadOnlyList<double> bounds = function.Items.TryGetValue("Bounds", out PdfObject? boundsObject)
            ? ReadNumberArray(boundsObject)
            : Array.Empty<double>();
        if (bounds.Count != functions.Items.Count - 1) return false;
        IReadOnlyList<double> domain = function.Items.TryGetValue("Domain", out PdfObject? domainObject)
            ? ReadNumberArray(domainObject)
            : Array.Empty<double>();
        if (!IsCanonicalUnitIntervals(domain, 1)) return false;
        double domainStart = domain[0];
        double domainEnd = domain[1];
        IReadOnlyList<double> encode = function.Items.TryGetValue("Encode", out PdfObject? encodeObject)
            ? ReadNumberArray(encodeObject)
            : Array.Empty<double>();
        if (!HasCanonicalFunctionEncode(encode, functions.Items.Count)) return false;
        if (function.Items.TryGetValue("Range", out PdfObject? rangeObject) &&
            !IsCanonicalUnitIntervals(ReadNumberArray(rangeObject), colorSpace.ComponentCount)) return false;
        if (bounds.Any(value => !IsFinite(value) || value <= domainStart || value >= domainEnd)) return false;
        for (int index = 1; index < bounds.Count; index++) if (bounds[index] <= bounds[index - 1]) return false;

        var result = new List<OfficeGradientStop>(functions.Items.Count + 1);
        PdfDictionary? first = ResolveFunctionDictionary(functions.Items[0]);
        if (!TryReadType2FunctionColors(first, colorSpace, IsFunctionReversed(encode, 0), out OfficeColor firstStart, out OfficeColor firstEnd)) return false;
        result.Add(new OfficeGradientStop(0D, firstStart));
        OfficeColor previousEnd = firstEnd;
        for (int i = 0; i < bounds.Count; i++) {
            double offset = Clamp01((bounds[i] - domainStart) / (domainEnd - domainStart));
            if (offset < result[result.Count - 1].Offset) return false;
            result.Add(new OfficeGradientStop(offset, previousEnd));
            PdfDictionary? next = ResolveFunctionDictionary(functions.Items[i + 1]);
            if (!TryReadType2FunctionColors(next, colorSpace, IsFunctionReversed(encode, i + 1), out OfficeColor nextStart, out OfficeColor nextEnd)) return false;
            if (nextStart != previousEnd) {
                result.Add(new OfficeGradientStop(offset, nextStart));
            }
            previousEnd = nextEnd;
        }
        result.Add(new OfficeGradientStop(1D, previousEnd));
        stops = result.AsReadOnly();
        return true;
    }

    private bool TryReadType2FunctionColors(PdfDictionary? function, PdfPageColorSpace colorSpace, bool reversed, out OfficeColor start, out OfficeColor end) {
        start = OfficeColor.Black;
        end = OfficeColor.Black;
        if (function == null || TryReadInteger(function.Items.TryGetValue("FunctionType", out PdfObject? type) ? type : null) != 2) return false;
        IReadOnlyList<double> domain = function.Items.TryGetValue("Domain", out PdfObject? domainObject)
            ? ReadNumberArray(domainObject)
            : Array.Empty<double>();
        if (!IsCanonicalUnitIntervals(domain, 1) ||
            !function.Items.TryGetValue("N", out PdfObject? exponentObject) ||
            ResolveObject(exponentObject) is not PdfNumber exponent ||
            !IsFinite(exponent.Value) || Math.Abs(exponent.Value - 1D) > 0.000000001D) return false;
        IReadOnlyList<double> c0 = function.Items.TryGetValue("C0", out PdfObject? c0Object)
            ? ReadNumberArray(c0Object)
            : new[] { 0D };
        IReadOnlyList<double> c1 = function.Items.TryGetValue("C1", out PdfObject? c1Object)
            ? ReadNumberArray(c1Object)
            : new[] { 1D };
        if (c0.Count != colorSpace.ComponentCount || c1.Count != colorSpace.ComponentCount ||
            !colorSpace.TryConvertColor(c0, out OfficeColor c0Color) ||
            !colorSpace.TryConvertColor(c1, out OfficeColor c1Color)) return false;
        if (function.Items.TryGetValue("Range", out PdfObject? rangeObject) &&
            !IsCanonicalUnitIntervals(ReadNumberArray(rangeObject), colorSpace.ComponentCount)) return false;
        start = reversed ? c1Color : c0Color;
        end = reversed ? c0Color : c1Color;
        return true;
    }

    private static bool HasValidFunctionIntervals(IReadOnlyList<double> values, int count) {
        if (values.Count != count * 2) return false;
        for (int index = 0; index < count; index++) {
            double minimum = values[index * 2];
            double maximum = values[index * 2 + 1];
            if (!IsFinite(minimum) || !IsFinite(maximum) || maximum <= minimum) return false;
        }
        return true;
    }

    private static bool IsCanonicalUnitIntervals(IReadOnlyList<double> values, int count) {
        if (!HasValidFunctionIntervals(values, count)) return false;
        for (int index = 0; index < count; index++) {
            if (Math.Abs(values[index * 2]) > 0.000000001D ||
                Math.Abs(values[(index * 2) + 1] - 1D) > 0.000000001D) return false;
        }
        return true;
    }

    private static bool HasCanonicalFunctionEncode(IReadOnlyList<double> values, int count) {
        if (!HasFiniteFunctionPairs(values, count)) return false;
        for (int index = 0; index < count; index++) {
            double first = values[index * 2];
            double second = values[(index * 2) + 1];
            bool forward = Math.Abs(first) <= 0.000000001D && Math.Abs(second - 1D) <= 0.000000001D;
            bool reverse = Math.Abs(first - 1D) <= 0.000000001D && Math.Abs(second) <= 0.000000001D;
            if (!forward && !reverse) return false;
        }
        return true;
    }

    private static bool HasFiniteFunctionPairs(IReadOnlyList<double> values, int count) {
        if (values.Count != count * 2) return false;
        for (int index = 0; index < values.Count; index++) if (!IsFinite(values[index])) return false;
        return true;
    }

    private static bool IsFunctionReversed(IReadOnlyList<double> encode, int functionIndex) {
        int offset = functionIndex * 2;
        return encode.Count > offset + 1 && encode[offset] > encode[offset + 1];
    }

    private PdfDictionary? ResolveFunctionDictionary(PdfObject? functionObject) {
        PdfObject? resolved = ResolveObject(functionObject);
        if (resolved is PdfArray array && array.Items.Count > 0) {
            resolved = ResolveObject(array.Items[0]);
        }

        return resolved is PdfDictionary dictionary ? dictionary : null;
    }

    private static byte ToColorByte(double value) =>
        (byte)Math.Round(Clamp01(value) * 255D);

    private static double Clamp01(double value) =>
        value < 0D ? 0D : value > 1D ? 1D : value;

    private Dictionary<string, PdfPageGraphicsStateResource> GetGraphicsStateResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageGraphicsStateResource>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("ExtGState", out PdfObject? extGStateObject)) {
            return result;
        }

        PdfDictionary? extGStates = ResolveDictionary(extGStateObject);
        if (extGStates == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in extGStates.Items) {
            PdfDictionary? state = ResolveDictionary(entry.Value);
            if (state == null) {
                continue;
            }

            double? fillOpacity = ReadOpacity(state, "ca");
            double? strokeOpacity = ReadOpacity(state, "CA");
            double? strokeWidth = ReadStrokeWidth(state);
            OfficeStrokeDashStyle? strokeDashStyle = ReadStrokeDashStyle(state);
            OfficeStrokeLineCap? strokeLineCap = ReadStrokeLineCap(state);
            OfficeStrokeLineJoin? strokeLineJoin = ReadStrokeLineJoin(state);
            OfficeBlendMode? blendMode = ReadBlendMode(state);
            bool hasUnsupportedBlendMode = state.Items.ContainsKey("BM") && !blendMode.HasValue;
            bool hasUnsupportedEntries = state.Items.Keys.Any(static key => key is not (
                "Type" or "ca" or "CA" or "LW" or "D" or "LC" or "LJ" or "BM" or "SMask"));
            bool hasSoftMask = state.Items.ContainsKey("SMask");
            PdfPageSoftMaskResource? softMask = hasSoftMask ? ReadSoftMask(state, resources) : null;
            bool clearsSoftMask = hasSoftMask &&
                state.Items.TryGetValue("SMask", out PdfObject? authoredSoftMask) &&
                ResolveEffectObject(authoredSoftMask) is PdfName { Name: "None" };
            if (fillOpacity.HasValue ||
                strokeOpacity.HasValue ||
                strokeWidth.HasValue ||
                strokeDashStyle.HasValue ||
                strokeLineCap.HasValue ||
                strokeLineJoin.HasValue ||
                blendMode.HasValue ||
                hasUnsupportedBlendMode ||
                hasSoftMask) {
                result[entry.Key] = new PdfPageGraphicsStateResource(fillOpacity, strokeOpacity, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin, blendMode, hasSoftMask, softMask, hasSoftMask && !clearsSoftMask && softMask == null, hasUnsupportedBlendMode, hasUnsupportedEntries);
            }
        }

        return result;
    }

    private Dictionary<string, PdfPageColorSpace> GetColorSpaceResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageColorSpace>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject)) {
            return result;
        }

        PdfDictionary? colorSpaces = ResolveDictionary(colorSpacesObject);
        if (colorSpaces == null) {
            return result;
        }
        if (colorSpaces.Items.Count > _limits.MaxColorSpaceResourcesPerPage) {
            throw new InvalidDataException(
                $"The page declares more than {_limits.MaxColorSpaceResourcesPerPage} color-space resources.");
        }

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (TryReadColorSpaceResource(entry.Value, out PdfPageColorSpace colorSpace)) {
                result[entry.Key] = colorSpace;
            }
        }

        return result;
    }

    private Dictionary<string, PdfPageColorSpace> GetPatternBaseColorSpaceResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageColorSpace>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            ResolveDictionary(colorSpacesObject) is not PdfDictionary colorSpaces) {
            return result;
        }
        if (colorSpaces.Items.Count > _limits.MaxColorSpaceResourcesPerPage) {
            throw new InvalidDataException(
                $"The page declares more than {_limits.MaxColorSpaceResourcesPerPage} color-space resources.");
        }

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (ResolveObject(entry.Value) is not PdfArray array ||
                array.Items.Count < 2 ||
                ResolveObject(array.Items[0]) is not PdfName { Name: "Pattern" } ||
                !TryReadColorSpaceResource(array.Items[1], out PdfPageColorSpace baseColorSpace) ||
                baseColorSpace == PdfPageColorSpaceKind.Pattern) {
                continue;
            }
            result[entry.Key] = baseColorSpace;
        }
        return result;
    }

    private bool TryReadColorSpaceResource(PdfObject? value, out PdfPageColorSpace colorSpace) {
        return TryReadExtendedColorSpaceResource(value, 0, out colorSpace);
    }

    private bool TryReadCalRgbColorSpace(PdfDictionary calibration, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (!calibration.Items.TryGetValue("WhitePoint", out PdfObject? whitePointObject)) return false;
        IReadOnlyList<double> whitePoint = ReadNumberArray(whitePointObject);
        if (whitePoint.Count != 3 || whitePoint.Any(static value => !IsFinite(value) || value <= 0D)) return false;

        IReadOnlyList<double>? gamma = null;
        if (calibration.Items.TryGetValue("Gamma", out PdfObject? gammaObject)) {
            gamma = ReadNumberArray(gammaObject);
            if (gamma.Count != 3 || gamma.Any(static value => !IsFinite(value) || value <= 0D)) return false;
        }

        IReadOnlyList<double>? matrix = null;
        if (calibration.Items.TryGetValue("Matrix", out PdfObject? matrixObject)) {
            matrix = ReadNumberArray(matrixObject);
            if (matrix.Count != 9 || matrix.Any(static value => !IsFinite(value))) return false;
        }

        colorSpace = PdfPageColorSpace.CalRgb(whitePoint[0], whitePoint[1], whitePoint[2], gamma, matrix);
        return true;
    }

    private static bool TryReadStandardColorSpaceName(string name, out PdfPageColorSpace colorSpace) {
        switch (name) {
            case "DeviceRGB":
            case "RGB":
                colorSpace = PdfPageColorSpaceKind.DeviceRgb;
                return true;
            case "DeviceCMYK":
            case "CMYK":
                colorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                return true;
            case "DeviceGray":
            case "G":
                colorSpace = PdfPageColorSpaceKind.DeviceGray;
                return true;
            case "CalGray":
                colorSpace = PdfPageColorSpaceKind.CalGray;
                return true;
            case "CalRGB":
                colorSpace = PdfPageColorSpaceKind.CalRgb;
                return true;
            case "Lab":
                colorSpace = PdfPageColorSpaceKind.Lab;
                return true;
            case "Pattern":
                colorSpace = PdfPageColorSpaceKind.Pattern;
                return true;
            default:
                colorSpace = PdfPageColorSpaceKind.DeviceGray;
                return false;
        }
    }

    private double? ReadOpacity(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(value) is not PdfNumber number) {
            return null;
        }

        if (number.Value < 0D) {
            return 0D;
        }

        return number.Value > 1D ? 1D : number.Value;
    }

    private double? ReadStrokeWidth(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("LW", out PdfObject? value) ||
            ResolveObject(value) is not PdfNumber number) {
            return null;
        }

        return Math.Max(0D, number.Value);
    }

    private OfficeStrokeDashStyle? ReadStrokeDashStyle(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("D", out PdfObject? value) ||
            ResolveObject(value) is not PdfArray dash ||
            dash.Items.Count == 0 ||
            ResolveObject(dash.Items[0]) is not PdfArray dashArray) {
            return null;
        }

        IReadOnlyList<double> values = ReadNumberArray(dashArray);
        if (values.Count == 0) {
            return OfficeStrokeDashStyle.Solid;
        }

        if (values.Count >= 6) {
            return OfficeStrokeDashStyle.DashDotDot;
        }

        if (values.Count >= 4) {
            return OfficeStrokeDashStyle.DashDot;
        }

        if (values.Count >= 2) {
            return values[0] <= values[1] ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Dash;
        }

        return OfficeStrokeDashStyle.Solid;
    }

    private OfficeStrokeLineCap? ReadStrokeLineCap(PdfDictionary dictionary) {
        int? lineCap = TryReadInteger(dictionary.Items.TryGetValue("LC", out PdfObject? value) ? value : null);
        switch (lineCap) {
            case 0:
                return OfficeStrokeLineCap.Butt;
            case 1:
                return OfficeStrokeLineCap.Round;
            case 2:
                return OfficeStrokeLineCap.Square;
            default:
                return null;
        }
    }

    private OfficeStrokeLineJoin? ReadStrokeLineJoin(PdfDictionary dictionary) {
        int? lineJoin = TryReadInteger(dictionary.Items.TryGetValue("LJ", out PdfObject? value) ? value : null);
        switch (lineJoin) {
            case 0:
                return OfficeStrokeLineJoin.Miter;
            case 1:
                return OfficeStrokeLineJoin.Round;
            case 2:
                return OfficeStrokeLineJoin.Bevel;
            default:
                return null;
        }
    }

    private void AddAnnotationAppearances(
        OfficeDrawing drawing,
        double pageHeight,
        Matrix2D pageTransform,
        TextContentParser.TextOutputBudget textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        if (!_pageDict.Items.TryGetValue("Annots", out PdfObject? annotationsObject)) {
            return;
        }

        PdfArray? annotations = ResolveArray(annotationsObject);
        if (annotations == null) {
            return;
        }
        EnsureAnnotationBudget(annotations);

        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        Dictionary<string, Func<byte[], int, string>> pageDecoders = ResourceResolver.GetBudgetedFontDecoders(_pageDict, _objects);
        Dictionary<string, Func<byte[], double>> pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        Dictionary<string, PdfFontResource> pageFonts = ResourceResolver.GetFontsForResources(pageResources, _objects);
        var activeForms = new HashSet<PdfStream>();
        for (int i = 0; i < annotations.Items.Count; i++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[i]);
            if (annotation == null ||
                !TryReadRectangle(annotation.Items.TryGetValue("Rect", out PdfObject? rectangleObject) ? rectangleObject : null, out (double X1, double Y1, double X2, double Y2) rectangle) ||
                IsHiddenAnnotation(annotation) ||
                !TryGetRenderableAnnotationAppearanceStream(annotation, out PdfStream appearanceStream, out _)) {
                continue;
            }

            PdfDictionary? appearanceResources = ResolveDictionary(appearanceStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? pageResources;
            string appearanceContent = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(appearanceStream)), appearanceStream.Dictionary);
            if (appearanceContent.Length == 0) {
                continue;
            }

            Matrix2D appearanceTransform = Matrix2D.Multiply(pageTransform, CreateAnnotationAppearanceTransform(rectangle, appearanceStream.Dictionary));
            var elements = new List<PdfPageDrawingElement>();
            var primitives = new List<PdfPageVisualPrimitive>();
            var renderedType3PaintOrders = new HashSet<double>();
            CollectVisualPrimitivesAndForms(
                appearanceContent,
                appearanceResources,
                appearanceTransform,
                drawing.Width,
                pageHeight,
                primitives.Add,
                activeForms,
                renderedType3PaintOrders: renderedType3PaintOrders,
                type3GlyphBudget: type3GlyphBudget,
                allowSupportedType3TransparencyGroups: true,
                type3ImageVisitor: (placement, image, effect) => elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect)),
                type3PrimitiveVisitor: (primitive, effect) => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect)),
                type3GroupVisitor: (group, transform, paintOrder, key, effect) => elements.Add(PdfPageDrawingElement.FromGroup(group, transform, paintOrder, key, elements.Count).WithEffect(effect)),
                textOutputBudget: textOutputBudget,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root);
            for (int primitiveIndex = 0; primitiveIndex < primitives.Count; primitiveIndex++) {
                elements.Add(PdfPageDrawingElement.FromPrimitive(primitives[primitiveIndex], elements.Count));
            }

            var textSpans = new List<PdfTextSpan>();
            Dictionary<string, Func<byte[], int, string>> appearanceDecoders = MergeDecoders(
                pageDecoders,
                ResourceResolver.GetBudgetedFontDecodersForForm(appearanceStream.Dictionary, _objects));
            Dictionary<string, Func<byte[], double>> appearanceWidthProviders = MergeWidthProviders(pageWidthProviders, ResourceResolver.GetFontWidthProviders(appearanceStream.Dictionary, _objects));
            Dictionary<string, PdfFontResource> appearanceFonts = MergeFonts(pageFonts, ResourceResolver.GetFontsForResources(appearanceResources, _objects));
            string transformedAppearanceContent = WrapContentWithTransform(appearanceContent, appearanceTransform, out int transformedAppearanceContentOffset);
            CollectTextAndForms(
                transformedAppearanceContent,
                appearanceResources,
                appearanceDecoders,
                appearanceWidthProviders,
                appearanceFonts,
                textSpans,
                activeForms,
                pageHeight,
                paintOrderOffset: -transformedAppearanceContentOffset,
                useLogicalTextFilters: false,
                textOutputBudget: textOutputBudget,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root,
                contentOrderOffset: -transformedAppearanceContentOffset);
            for (int textIndex = 0; textIndex < textSpans.Count; textIndex++) {
                if (renderedType3PaintOrders.Contains(textSpans[textIndex].PaintOrder)) continue;
                elements.Add(PdfPageDrawingElement.FromText(textSpans[textIndex], elements.Count));
            }

            var imagePlacements = new List<PdfImagePlacement>();
            CollectImagePlacementsAndForms(appearanceContent, appearanceResources, 0, appearanceTransform, pageHeight, imagePlacements, activeForms, pageContentBudget: pageContentBudget);
            if (imagePlacements.Count > 0) {
                IReadOnlyList<PdfExtractedImage> images = GetImagesForResources(appearanceResources, 0, imagePlacements, colorizeImageMasks: true);
                for (int imageIndex = 0; imageIndex < imagePlacements.Count; imageIndex++) {
                    PdfImagePlacement placement = imagePlacements[imageIndex];
                    PdfExtractedImage? image = FindImage(images, placement);
                    if (image != null) {
                        elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count));
                    }
                }
            }
            var appearanceEffects = new List<PdfPageDrawingEffectTransition>();
            CollectGraphicsEffectTransitions(
                appearanceContent,
                appearanceResources,
                appearanceTransform,
                pageHeight,
                appearanceEffects,
                new HashSet<PdfStream>(),
                PdfPageDrawingEffect.Default,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root);
            SortGraphicsEffectTransitions(appearanceEffects);
            OverlayDrawingEffects(elements, appearanceEffects);

            SortDrawingElements(elements);
            var appearanceSoftMasks = new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
            var activeAppearanceSoftMasks = new HashSet<PdfStream>();
            for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
                AddDrawingElement(drawing, pageHeight, appearanceTransform, elements[elementIndex], appearanceSoftMasks, activeAppearanceSoftMasks, textOutputBudget, pageContentBudget, type3GlyphBudget);
            }
        }
    }

    private bool TryGetRenderableAnnotationAppearanceStream(
        PdfDictionary annotation,
        out PdfStream stream,
        out bool synthesized) {
        if (TryGetNormalAppearanceStream(annotation, out stream)) {
            synthesized = false;
            return true;
        }

        synthesized = PdfAnnotationFlattener.TryCreateSyntheticAppearanceStream(_objects, annotation, out stream);
        return synthesized;
    }

    private bool TryGetNormalAppearanceStream(PdfDictionary annotation, out PdfStream stream) {
        stream = null!;
        PdfDictionary? appearance = ResolveDictionary(annotation.Items.TryGetValue("AP", out PdfObject? appearanceObject) ? appearanceObject : null);
        if (appearance == null || !appearance.Items.TryGetValue("N", out PdfObject? normalAppearanceObject)) {
            return false;
        }

        PdfObject? normalAppearance = ResolveObject(normalAppearanceObject);
        if (normalAppearance is PdfStream directStream) {
            stream = directStream;
            return true;
        }

        if (normalAppearance is not PdfDictionary stateDictionary || stateDictionary.Items.Count == 0) {
            return false;
        }

        if (stateDictionary.Items.Count > _limits.MaxFormFieldAppearanceStates) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.FormAppearanceStates,
                _limits.MaxFormFieldAppearanceStates,
                stateDictionary.Items.Count);
        }

        if (annotation.Items.TryGetValue("AS", out PdfObject? appearanceStateObject) &&
            ResolveObject(appearanceStateObject) is PdfName appearanceState) {
            if (stateDictionary.Items.TryGetValue(appearanceState.Name, out PdfObject? stateObject) &&
                ResolveObject(stateObject) is PdfStream stateStream) {
                stream = stateStream;
                return true;
            }

            return false;
        }

        foreach (KeyValuePair<string, PdfObject> state in stateDictionary.Items) {
            if (string.Equals(state.Key, "Off", StringComparison.Ordinal)) {
                continue;
            }

            if (ResolveObject(state.Value) is PdfStream fallbackStream) {
                stream = fallbackStream;
                return true;
            }
        }

        foreach (KeyValuePair<string, PdfObject> state in stateDictionary.Items) {
            if (ResolveObject(state.Value) is PdfStream fallbackStream) {
                stream = fallbackStream;
                return true;
            }
        }

        return false;
    }

    private Matrix2D CreateAnnotationAppearanceTransform((double X1, double Y1, double X2, double Y2) rectangle, PdfDictionary appearanceDictionary) {
        double bboxX1 = 0D;
        double bboxY1 = 0D;
        double bboxWidth = rectangle.X2 - rectangle.X1;
        double bboxHeight = rectangle.Y2 - rectangle.Y1;
        if (TryReadBox(appearanceDictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null, out (double X1, double Y1, double X2, double Y2) bbox)) {
            bboxX1 = bbox.X1;
            bboxY1 = bbox.Y1;
            bboxWidth = bbox.X2 - bbox.X1;
            bboxHeight = bbox.Y2 - bbox.Y1;
        }

        double scaleX = bboxWidth > 0D ? (rectangle.X2 - rectangle.X1) / bboxWidth : 1D;
        double scaleY = bboxHeight > 0D ? (rectangle.Y2 - rectangle.Y1) / bboxHeight : 1D;
        var rectangleTransform = new Matrix2D(
            scaleX,
            0D,
            0D,
            scaleY,
            rectangle.X1 - (bboxX1 * scaleX),
            rectangle.Y1 - (bboxY1 * scaleY));
        return Matrix2D.Multiply(rectangleTransform, ReadAppearanceMatrix(appearanceDictionary));
    }

    private Matrix2D ReadAppearanceMatrix(PdfDictionary appearanceDictionary) {
        if (!appearanceDictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject) ||
            ResolveObject(matrixObject) is not PdfArray matrix ||
            matrix.Items.Count < 6) {
            return Matrix2D.Identity;
        }

        return new Matrix2D(
            ReadMatrixNumber(matrix, 0, 1D),
            ReadMatrixNumber(matrix, 1, 0D),
            ReadMatrixNumber(matrix, 2, 0D),
            ReadMatrixNumber(matrix, 3, 1D),
            ReadMatrixNumber(matrix, 4, 0D),
            ReadMatrixNumber(matrix, 5, 0D));
    }

    private double ReadMatrixNumber(PdfArray matrix, int index, double fallback) =>
        ResolveObject(matrix.Items[index]) is PdfNumber number ? number.Value : fallback;

    private bool TryReadBox(PdfObject? obj, out (double X1, double Y1, double X2, double Y2) box) =>
        TryReadRectangle(obj, out box);

    private bool IsHiddenAnnotation(PdfDictionary annotation) {
        int? flags = TryReadInteger(annotation.Items.TryGetValue("F", out PdfObject? flagsObject) ? flagsObject : null);
        if (!flags.HasValue) {
            return false;
        }

        const int invisible = 1;
        const int hidden = 2;
        const int noView = 32;
        return (flags.Value & (invisible | hidden | noView)) != 0;
    }

    private void AddTextSpans(OfficeDrawing drawing, double pageHeight, Matrix2D pageTransform) {
        IReadOnlyList<PdfTextSpan> spans = GetVisualTextSpans(pageHeight, pageTransform);
        for (int i = 0; i < spans.Count; i++) {
            AddTextSpan(drawing, pageHeight, spans[i]);
        }
    }

    private IReadOnlyList<PdfTextSpan> GetVisualTextSpans(
        double pageHeight,
        Matrix2D pageTransform,
        TextContentParser.TextOutputBudget? textOutputBudget = null,
        PageContentBudget? pageContentBudget = null) {
        textOutputBudget ??= CreateTextOutputBudget();
        pageContentBudget ??= new PageContentBudget(this);
        var spans = new List<PdfTextSpan>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        Dictionary<string, Func<byte[], int, string>> pageDecoders = ResourceResolver.GetBudgetedFontDecoders(_pageDict, _objects);
        Dictionary<string, Func<byte[], double>> pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        Dictionary<string, PdfFontResource> pageFonts = ResourceResolver.GetFontsForResources(pageResources, _objects);
        var activeForms = new HashSet<PdfStream>();

        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            string transformedContent = WrapContentWithTransform(content, pageTransform, out int transformedContentOffset);
            CollectTextAndForms(
                transformedContent,
                pageResources,
                pageDecoders,
                pageWidthProviders,
                pageFonts,
                spans,
                activeForms,
                pageHeight,
                paintOrderOffset: -transformedContentOffset,
                useLogicalTextFilters: false,
                textOutputBudget: textOutputBudget,
                pageContentBudget: pageContentBudget,
                contentOrderPrefix: PdfContentOrderKey.Root,
                contentOrderOffset: -transformedContentOffset);
        }

        return spans.Count == 0 ? Array.Empty<PdfTextSpan>() : spans.AsReadOnly();
    }

    private TextContentParser.TextOutputBudget CreateTextOutputBudget() =>
        new(_limits.MaxActualTextCharacters, _limits.MaxDecodedTextCharacters);

    private static void AddTextSpan(OfficeDrawing drawing, double pageHeight, PdfTextSpan span) {
        if (string.IsNullOrEmpty(span.Text) || !span.IsVisible) {
            return;
        }

        double height = Math.Max(1D, span.FontSize * 1.25D);
        double width = Math.Max(span.Advance, span.Text.Length * span.FontSize * 0.55D);
        double rawX = span.X;
        double rawY = pageHeight - span.Y - span.FontSize;
        if (!HasVisibleOverlap(rawX, rawY, width, height, drawing.Width, drawing.Height)) {
            return;
        }

        double x = rawX;
        double y = rawY;
        double clippedRight = Math.Min(rawX + width, drawing.Width);
        double clippedBottom = Math.Min(rawY + height, drawing.Height);
        double baselineY = pageHeight - span.Y;
        if (!span.ClipPath.HasValue &&
            (rawX < 0D || rawY < 0D || rawX + width > drawing.Width || rawY + height > drawing.Height)) {
            PdfPageClipPath pageClip = PdfPageClipPath.Rectangle(0D, 0D, drawing.Width, drawing.Height);
            if (TryAddClippedTextSpan(drawing, span, x, y, width, height, baselineY, pageClip)) {
                return;
            }
        }

        x = Clamp(rawX, 0D, drawing.Width);
        y = Clamp(rawY, 0D, drawing.Height);
        baselineY = Clamp(baselineY, 0D, drawing.Height);
        width = Math.Max(1D, clippedRight - x);
        height = Math.Max(1D, clippedBottom - y);
        if (TryAddClippedTextSpan(drawing, span, x, y, width, height, baselineY)) {
            return;
        }

        drawing.AddText(
            span.Text,
            x,
            y,
            width,
            height,
            ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily),
            span.Color ?? OfficeColor.Black,
            rotationDegrees: -span.RotationDegrees,
            rotationCenterX: x,
            rotationCenterY: baselineY,
            wrapText: false);
    }

    private static bool TryAddClippedTextSpan(OfficeDrawing drawing, PdfTextSpan span, double x, double y, double width, double height, double baselineY, PdfPageClipPath? overrideClipPath = null) {
        PdfPageClipPath? activeClipPath = overrideClipPath ?? span.ClipPath;
        if (!activeClipPath.HasValue) {
            return false;
        }

        PdfPageClipPath clip = activeClipPath.Value;
        if (clip.Width <= 0D || clip.Height <= 0D) {
            return true;
        }

        OfficeClipPath? officeClipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (officeClipPath == null) {
            return false;
        }

        double clipRight = clip.X + clip.Width;
        double clipBottom = clip.Y + clip.Height;
        if (clip.IsRectangle && x >= clip.X && y >= clip.Y && x + width <= clipRight && y + height <= clipBottom) {
            return false;
        }

        if (x + width <= clip.X || y + height <= clip.Y || x >= clipRight || y >= clipBottom) {
            return true;
        }

        double localX = x - clip.X;
        double localY = y - clip.Y;
        if (clip.X < 0D ||
            clip.Y < 0D ||
            clipRight > drawing.Width ||
            clipBottom > drawing.Height) {
            if (!TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out PdfPageClipPath drawingClip)) {
                return true;
            }

            clip = drawingClip;
            officeClipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
            if (officeClipPath == null) {
                return false;
            }

            localX = x - clip.X;
            localY = y - clip.Y;
        }

        double textWidth = Math.Max(1D, width);
        double textHeight = Math.Max(1D, height);
        drawing.AddClippedText(
            span.Text,
            x,
            y,
            textWidth,
            textHeight,
            clip.X,
            clip.Y,
            officeClipPath,
            ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily),
            span.Color ?? OfficeColor.Black,
            rotationDegrees: -span.RotationDegrees,
            rotationCenterX: x,
            rotationCenterY: baselineY,
            wrapText: false);
        return true;
    }

    private static OfficeFontInfo ToOfficeFontInfo(string? baseFont, double size, string? drawingFontFamily = null) {
        string normalized = StripSubsetPrefix(baseFont);
        OfficeFontStyle style = OfficeFontStyle.Regular;
        if (ContainsFontStyleToken(normalized, "Bold") ||
            ContainsFontStyleToken(normalized, "Black") ||
            ContainsFontStyleToken(normalized, "Heavy") ||
            ContainsFontStyleToken(normalized, "Demi") ||
            ContainsFontStyleToken(normalized, "SemiBold")) {
            style |= OfficeFontStyle.Bold;
        }

        if (ContainsFontStyleToken(normalized, "Italic") ||
            ContainsFontStyleToken(normalized, "Oblique")) {
            style |= OfficeFontStyle.Italic;
        }

        string family = string.IsNullOrWhiteSpace(drawingFontFamily)
            ? ResolveOfficeFontFamily(normalized)
            : drawingFontFamily!;
        return new OfficeFontInfo(family, size, style);
    }

    private static string ResolveOfficeFontFamily(string baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont)) {
            return "Helvetica";
        }

        string normalized = baseFont.Replace('_', ' ');
        if (normalized.StartsWith("Times-", StringComparison.Ordinal) ||
            normalized.StartsWith("TimesNewRoman", StringComparison.OrdinalIgnoreCase) ||
            normalized.StartsWith("Times New Roman", StringComparison.OrdinalIgnoreCase)) {
            return "Times New Roman";
        }

        if (normalized.StartsWith("Courier", StringComparison.OrdinalIgnoreCase)) {
            return "Courier New";
        }

        if (normalized.StartsWith("Helvetica", StringComparison.OrdinalIgnoreCase)) {
            return "Helvetica";
        }

        int hyphen = normalized.IndexOf('-');
        if (hyphen > 0) {
            normalized = normalized.Substring(0, hyphen);
        }

        normalized = RemoveFontSuffix(normalized, "BoldItalic");
        normalized = RemoveFontSuffix(normalized, "BoldOblique");
        normalized = RemoveFontSuffix(normalized, "SemiBold");
        normalized = RemoveFontSuffix(normalized, "DemiBold");
        normalized = RemoveFontSuffix(normalized, "Bold");
        normalized = RemoveFontSuffix(normalized, "Italic");
        normalized = RemoveFontSuffix(normalized, "Oblique");
        normalized = RemoveFontSuffix(normalized, "Regular");
        normalized = RemoveFontSuffix(normalized, "PSMT");
        normalized = RemoveFontSuffix(normalized, "MT");
        return string.IsNullOrWhiteSpace(normalized) ? "Helvetica" : normalized.Trim();
    }

    private static string StripSubsetPrefix(string? baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont)) {
            return string.Empty;
        }

        string value = baseFont!.Trim();
        if (value.Length > 7 && value[6] == '+') {
            for (int i = 0; i < 6; i++) {
                char ch = value[i];
                if (ch < 'A' || ch > 'Z') {
                    return value;
                }
            }

            return value.Substring(7);
        }

        return value;
    }

    private static bool ContainsFontStyleToken(string fontName, string token) =>
        System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(fontName, token, System.Globalization.CompareOptions.IgnoreCase) >= 0;

    private static string RemoveFontSuffix(string value, string suffix) =>
        value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            ? value.Substring(0, value.Length - suffix.Length)
            : value;

    private void AddImages(OfficeDrawing drawing, double pageHeight, Matrix2D pageTransform) {
        IReadOnlyList<PdfImagePlacement> placements = GetVisualImagePlacements(pageHeight, pageTransform);
        if (placements.Count == 0) {
            return;
        }

        IReadOnlyList<PdfExtractedImage> images = GetImages(0, placements, colorizeImageMasks: true);
        AddImagePlacements(drawing, pageHeight, placements, images);
    }

    private IReadOnlyList<PdfImagePlacement> GetVisualImagePlacements(double pageHeight, Matrix2D pageTransform, PageContentBudget? pageContentBudget = null) {
        pageContentBudget ??= new PageContentBudget(this);
        var placements = new List<PdfImagePlacement>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();

        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            CollectImagePlacementsAndForms(
                content,
                pageResources,
                0,
                pageTransform,
                pageHeight,
                placements,
                activeForms,
                pageContentBudget: pageContentBudget);
        }

        return placements.Count == 0 ? Array.Empty<PdfImagePlacement>() : placements.AsReadOnly();
    }

    private static void AddImagePlacements(OfficeDrawing drawing, double pageHeight, IReadOnlyList<PdfImagePlacement> placements, IReadOnlyList<PdfExtractedImage> images) {
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            PdfExtractedImage? image = FindImage(images, placement);
            if (image == null) {
                continue;
            }

            AddImagePlacement(drawing, pageHeight, placement, image);
        }
    }

    private static void AddImagePlacement(OfficeDrawing drawing, double pageHeight, PdfImagePlacement placement, PdfExtractedImage image) {
        if (!image.IsImageFile || placement.Width <= 0D || placement.Height <= 0D) {
            return;
        }

        if (!TryCreateImageProjection(placement, pageHeight, drawing.Width, drawing.Height, out OfficeImageProjection projection)) {
            return;
        }

        if (TryAddClippedImagePlacement(drawing, placement, image, projection)) {
            return;
        }

        drawing.AddImage(image.Bytes, image.MimeType, projection, opacity: placement.ImageOpacity ?? 1D);
    }

    private static bool TryAddClippedImagePlacement(OfficeDrawing drawing, PdfImagePlacement placement, PdfExtractedImage image, OfficeImageProjection projection) {
        PdfPageClipPath? activeClipPath = placement.ClipPath;
        if (!activeClipPath.HasValue && projection.HasTransform) {
            (double left, double top, double right, double bottom) = projection.GetDestinationBounds();
            if (HasPositiveArea(left, top, right - left, bottom - top, drawing.Width, drawing.Height)) {
                return false;
            }

            if (!TryCreatePageClip(left, top, right, bottom, drawing.Width, drawing.Height, out PdfPageClipPath pageClip)) {
                return true;
            }

            activeClipPath = pageClip;
        }

        if (!activeClipPath.HasValue || (activeClipPath.Value.IsRectangle && !projection.HasTransform)) {
            // Plain axis-aligned rectangle clips are converted into source crops by TryCreateImageProjection.
            return false;
        }

        PdfPageClipPath clip = activeClipPath.Value;
        if (clip.Width <= 0D || clip.Height <= 0D) {
            return true;
        }

        if (!TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out PdfPageClipPath drawingClip)) {
            return true;
        }

        clip = drawingClip;
        OfficeClipPath? clipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (clipPath == null) {
            return false;
        }

        drawing.AddClippedImage(image.Bytes, image.MimeType, projection, clip.X, clip.Y, clipPath, opacity: placement.ImageOpacity ?? 1D);
        return true;
    }

    private static bool TryCreatePageClip(double left, double top, double right, double bottom, double drawingWidth, double drawingHeight, out PdfPageClipPath clip) {
        double visibleLeft = Math.Max(0D, left);
        double visibleTop = Math.Max(0D, top);
        double visibleRight = Math.Min(drawingWidth, right);
        double visibleBottom = Math.Min(drawingHeight, bottom);
        double visibleWidth = visibleRight - visibleLeft;
        double visibleHeight = visibleBottom - visibleTop;
        if (visibleWidth <= 0D || visibleHeight <= 0D) {
            clip = default;
            return false;
        }

        clip = PdfPageClipPath.Rectangle(visibleLeft, visibleTop, visibleWidth, visibleHeight);
        return true;
    }

    private static bool TryFitClipToDrawing(PdfPageClipPath clip, double drawingWidth, double drawingHeight, out PdfPageClipPath drawingClip) {
        if (clip.X >= 0D &&
            clip.Y >= 0D &&
            clip.X + clip.Width <= drawingWidth &&
            clip.Y + clip.Height <= drawingHeight) {
            drawingClip = clip;
            return true;
        }

        PdfPageClipPath pageClip = PdfPageClipPath.Rectangle(0D, 0D, drawingWidth, drawingHeight);
        if (!IntersectClipBounds(clip, pageClip, out PdfPageClipPath intersection)) {
            drawingClip = default;
            return false;
        }

        drawingClip = clip.WithBounds(intersection);
        return true;
    }

    private static bool TryCreateImageProjection(
        PdfImagePlacement placement,
        double pageHeight,
        double drawingWidth,
        double drawingHeight,
        out OfficeImageProjection projection,
        bool allowAxisAlignedFallback = true) {
        if (!IsPlainAxisAlignedImagePlacement(placement)) {
            if (TryCreateTransformedImageProjection(placement, pageHeight, drawingWidth, drawingHeight, out projection)) {
                return true;
            }
            if (!allowAxisAlignedFallback) return false;
        }

        double imageX = placement.X;
        double imageY = pageHeight - placement.Y - placement.Height;
        projection = default;

        if (placement.Width <= 0D || placement.Height <= 0D) {
            return false;
        }

        PdfPageClipPath? clip = placement.ClipPath;
        double imageRight = imageX + placement.Width;
        double imageBottom = imageY + placement.Height;
        double visibleLeft = Math.Max(imageX, 0D);
        double visibleTop = Math.Max(imageY, 0D);
        double visibleRight = Math.Min(imageRight, drawingWidth);
        double visibleBottom = Math.Min(imageBottom, drawingHeight);

        if (!clip.HasValue || !clip.Value.IsRectangle) {
            double pageVisibleWidth = visibleRight - visibleLeft;
            double pageVisibleHeight = visibleBottom - visibleTop;
            if (pageVisibleWidth <= 0D || pageVisibleHeight <= 0D) {
                return false;
            }

            if (NearlyEqual(visibleLeft, imageX) &&
                NearlyEqual(visibleTop, imageY) &&
                NearlyEqual(pageVisibleWidth, placement.Width) &&
                NearlyEqual(pageVisibleHeight, placement.Height)) {
                // Normalize sub-point producer rounding at the page boundary. Keeping
                // the original near-equal coordinates can place a valid full-page
                // image microscopically outside the Drawing contract.
                projection = new OfficeImageProjection(new OfficeImagePlacement(
                    visibleLeft,
                    visibleTop,
                    pageVisibleWidth,
                    pageVisibleHeight));
                return true;
            }

            var pageCrop = OfficeImageSourceCrop.FromClampedFractions(
                (visibleLeft - imageX) / placement.Width,
                (visibleTop - imageY) / placement.Height,
                (imageRight - visibleRight) / placement.Width,
                (imageBottom - visibleBottom) / placement.Height);
            projection = new OfficeImageProjection(new OfficeImagePlacement(visibleLeft, visibleTop, pageVisibleWidth, pageVisibleHeight), pageCrop);
            return true;
        }

        double clipLeft = clip.Value.X;
        double clipTop = clip.Value.Y;
        double clipRight = clipLeft + clip.Value.Width;
        double clipBottom = clipTop + clip.Value.Height;
        visibleLeft = Math.Max(visibleLeft, clipLeft);
        visibleTop = Math.Max(visibleTop, clipTop);
        visibleRight = Math.Min(visibleRight, clipRight);
        visibleBottom = Math.Min(visibleBottom, clipBottom);
        double visibleWidth = visibleRight - visibleLeft;
        double visibleHeight = visibleBottom - visibleTop;
        if (visibleWidth <= 0D || visibleHeight <= 0D ||
            !HasPositiveArea(visibleLeft, visibleTop, visibleWidth, visibleHeight, drawingWidth, drawingHeight)) {
            return false;
        }

        var crop = OfficeImageSourceCrop.FromClampedFractions(
            (visibleLeft - imageX) / placement.Width,
            (visibleTop - imageY) / placement.Height,
            (imageRight - visibleRight) / placement.Width,
            (imageBottom - visibleBottom) / placement.Height);
        projection = new OfficeImageProjection(new OfficeImagePlacement(visibleLeft, visibleTop, visibleWidth, visibleHeight), crop);
        return true;
    }

    private static bool IsPlainAxisAlignedImagePlacement(PdfImagePlacement placement) =>
        NearlyEqual(placement.B, 0D) &&
        NearlyEqual(placement.C, 0D) &&
        placement.A >= 0D &&
        placement.D >= 0D;

    private static bool TryCreateTransformedImageProjection(PdfImagePlacement placement, double pageHeight, double drawingWidth, double drawingHeight, out OfficeImageProjection projection) {
        projection = default;
        double m11 = placement.A;
        double m12 = -placement.B;
        double m21 = -placement.C;
        double m22 = placement.D;
        double offsetX = placement.C + placement.E;
        double offsetY = pageHeight - placement.D - placement.F;
        double width = Math.Sqrt((m11 * m11) + (m12 * m12));
        double height = Math.Sqrt((m21 * m21) + (m22 * m22));
        if (width <= 0D || height <= 0D) {
            return false;
        }

        double dot = (m11 * m21) + (m12 * m22);
        if (!NearlyEqual(dot, 0D)) {
            return false;
        }

        return TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: false, flipVertical: false, drawingWidth, drawingHeight, out projection) ||
               TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: true, flipVertical: false, drawingWidth, drawingHeight, out projection) ||
               TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: false, flipVertical: true, drawingWidth, drawingHeight, out projection) ||
               TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: true, flipVertical: true, drawingWidth, drawingHeight, out projection);
    }

    private static bool TryCreateImageProjectionCandidate(
        double m11,
        double m12,
        double m21,
        double m22,
        double offsetX,
        double offsetY,
        double width,
        double height,
        bool flipHorizontal,
        bool flipVertical,
        double drawingWidth,
        double drawingHeight,
        out OfficeImageProjection projection) {
        projection = default;
        double columnSign = flipHorizontal ? -1D : 1D;
        double rowSign = flipVertical ? -1D : 1D;
        double cos = m11 / (columnSign * width);
        double sin = m12 / (columnSign * width);
        double baseColumnX = width * cos;
        double baseColumnY = width * sin;
        double baseRowX = -height * sin;
        double baseRowY = height * cos;
        if (!NearlyEqual(m21, rowSign * baseRowX) ||
            !NearlyEqual(m22, rowSign * baseRowY)) {
            return false;
        }

        double unflippedOffsetX = offsetX;
        double unflippedOffsetY = offsetY;
        if (flipHorizontal) {
            unflippedOffsetX -= baseColumnX;
            unflippedOffsetY -= baseColumnY;
        }

        if (flipVertical) {
            unflippedOffsetX -= baseRowX;
            unflippedOffsetY -= baseRowY;
        }

        double x = unflippedOffsetX - (width / 2D) + (cos * width / 2D) - (sin * height / 2D);
        double y = unflippedOffsetY - (height / 2D) + (sin * width / 2D) + (cos * height / 2D);
        if (!IsFinite(x) || !IsFinite(y)) {
            return false;
        }

        double rotationDegrees = Math.Atan2(sin, cos) * 180D / Math.PI;
        projection = new OfficeImageProjection(
            new OfficeImagePlacement(x, y, width, height),
            rotationDegrees: rotationDegrees,
            flipHorizontal: flipHorizontal,
            flipVertical: flipVertical);
        (double left, double top, double right, double bottom) = projection.GetDestinationBounds();
        return HasVisibleOverlap(left, top, right - left, bottom - top, drawingWidth, drawingHeight);
    }

    private static PdfExtractedImage? FindImage(IReadOnlyList<PdfExtractedImage> images, PdfImagePlacement placement) {
        for (int i = 0; i < images.Count; i++) {
            PdfExtractedImage image = images[i];
            if (string.Equals(image.ResourceName, placement.ResourceName, StringComparison.Ordinal) &&
                image.ObjectNumber == placement.ObjectNumber &&
                image.DirectStreamIdentity == placement.DirectStreamIdentity &&
                (!image.IsImageMask || image.ImageMaskColor.Equals(placement.ImageMaskColor))) {
                return image;
            }
        }

        return null;
    }

    private static bool HasPositiveArea(double x, double y, double width, double height, double maxWidth, double maxHeight) =>
        width > 0D &&
        height > 0D &&
        x >= 0D &&
        y >= 0D &&
        x + width <= maxWidth &&
        y + height <= maxHeight;

    private static bool HasVisibleOverlap(double x, double y, double width, double height, double maxWidth, double maxHeight) =>
        IsFinite(x) &&
        IsFinite(y) &&
        IsFinite(width) &&
        IsFinite(height) &&
        IsFinite(maxWidth) &&
        IsFinite(maxHeight) &&
        width > 0D &&
        height > 0D &&
        x < maxWidth &&
        y < maxHeight &&
        x + width > 0D &&
        y + height > 0D;

    private static double Clamp(double value, double min, double max) {
        if (value < min) {
            return min;
        }

        return value > max ? max : value;
    }

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;

    private enum PdfPageDrawingElementKind {
        Primitive,
        Text,
        Image,
        Group
    }

    private readonly struct PdfPageDrawingElement {
        private PdfPageDrawingElement(
            PdfPageDrawingElementKind kind,
            double paintOrder,
            int sequence,
            PdfPageVisualPrimitive primitive,
            PdfTextSpan? textSpan,
            PdfImagePlacement? imagePlacement,
            PdfExtractedImage? image,
            OfficeDrawing? groupDrawing,
            OfficeTransform groupTransform,
            PdfContentOrderKey? groupContentOrderKey,
            PdfPageDrawingEffect effect) {
            Kind = kind;
            PaintOrder = paintOrder;
            Sequence = sequence;
            Primitive = primitive;
            TextSpan = textSpan;
            ImagePlacement = imagePlacement;
            Image = image;
            GroupDrawing = groupDrawing;
            GroupTransform = groupTransform;
            GroupContentOrderKey = groupContentOrderKey;
            Effect = effect;
        }

        public static PdfPageDrawingElement FromPrimitive(PdfPageVisualPrimitive primitive, int sequence) =>
            new PdfPageDrawingElement(PdfPageDrawingElementKind.Primitive, primitive.PaintOrder, sequence, primitive, null, null, null, null, OfficeTransform.Identity, null, PdfPageDrawingEffect.Default);

        public static PdfPageDrawingElement FromText(PdfTextSpan textSpan, int sequence) =>
            new PdfPageDrawingElement(PdfPageDrawingElementKind.Text, textSpan.PaintOrder, sequence, default, textSpan, null, null, null, OfficeTransform.Identity, null, PdfPageDrawingEffect.Default);

        public static PdfPageDrawingElement FromImage(PdfImagePlacement imagePlacement, PdfExtractedImage image, int sequence) =>
            new PdfPageDrawingElement(PdfPageDrawingElementKind.Image, imagePlacement.PaintOrder, sequence, default, null, imagePlacement, image, null, OfficeTransform.Identity, null, PdfPageDrawingEffect.Default);

        public static PdfPageDrawingElement FromGroup(OfficeDrawing groupDrawing, double paintOrder, PdfContentOrderKey? contentOrderKey, int sequence) =>
            FromGroup(groupDrawing, OfficeTransform.Identity, paintOrder, contentOrderKey, sequence);

        public static PdfPageDrawingElement FromGroup(OfficeDrawing groupDrawing, OfficeTransform transform, double paintOrder, PdfContentOrderKey? contentOrderKey, int sequence) =>
            new PdfPageDrawingElement(PdfPageDrawingElementKind.Group, paintOrder, sequence, default, null, null, null, groupDrawing, transform, contentOrderKey, PdfPageDrawingEffect.Default);

        public PdfPageDrawingElementKind Kind { get; }

        public double PaintOrder { get; }

        public int Sequence { get; }

        public PdfPageVisualPrimitive Primitive { get; }

        public PdfTextSpan? TextSpan { get; }

        public PdfImagePlacement? ImagePlacement { get; }

        public PdfExtractedImage? Image { get; }

        public OfficeDrawing? GroupDrawing { get; }

        public OfficeTransform GroupTransform { get; }

        private PdfContentOrderKey? GroupContentOrderKey { get; }

        public PdfPageDrawingEffect Effect { get; }

        internal PdfContentOrderKey? ContentOrderKey => Kind switch {
            PdfPageDrawingElementKind.Primitive => Primitive.ContentOrderKey,
            PdfPageDrawingElementKind.Text => TextSpan?.ContentOrderKey,
            PdfPageDrawingElementKind.Image => ImagePlacement?.ContentOrderKey,
            PdfPageDrawingElementKind.Group => GroupContentOrderKey,
            _ => null
        };

        public PdfPageDrawingElement WithEffect(PdfPageDrawingEffect effect) =>
            new PdfPageDrawingElement(Kind, PaintOrder, Sequence, Primitive, TextSpan, ImagePlacement, Image, GroupDrawing, GroupTransform, GroupContentOrderKey, effect);
    }

}
