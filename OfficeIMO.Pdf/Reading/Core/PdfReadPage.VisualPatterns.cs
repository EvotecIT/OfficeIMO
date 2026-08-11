using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static void AddTilingPatternFill(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        PdfPageTilingPatternPaint paint = primitive.FillTilingPattern!;
        if (primitive.Width <= 0D || primitive.Height <= 0D) return;
        PdfPageClipPath shapeClip;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
            shapeClip = PdfPageClipPath.Rectangle(primitive.X, primitive.Y, primitive.Width, primitive.Height);
        } else if (!PdfPageClipPath.TryCreatePath(primitive.PathCommands, primitive.FillRule, out shapeClip)) {
            return;
        }

        if (primitive.ClipPath.HasValue) {
            shapeClip = PdfPageClipPath.ResolveActiveClip(primitive.ClipPath.Value, shapeClip);
        }
        if (!TryFitClipToDrawing(shapeClip, drawing.Width, drawing.Height, out PdfPageClipPath fitted)) return;
        OfficeClipPath? clip = fitted.ToOfficeClipPath(fitted.X, fitted.Y);
        if (clip == null) return;

        OfficeDrawing tile = paint.Resource.Tile.Clone();
        if (paint.Tint.HasValue) tile.ApplyColorTint(paint.Tint.Value);
        var patternDrawing = new OfficeDrawing(fitted.Width, fitted.Height);
        OfficeTransform localTransform = paint.Transform.Then(OfficeTransform.Translate(-fitted.X, -fitted.Y));
        patternDrawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, fitted.Width, fitted.Height),
            paint.Resource.HorizontalStep,
            paint.Resource.VerticalStep,
            localTransform,
            maximumTileCount: 16384,
            opacity: paint.Opacity);
        drawing.AddClippedDrawing(patternDrawing, fitted.X, fitted.Y, clip);
    }

    private static void AddTilingPatternStroke(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        PdfPageTilingPatternPaint paint = primitive.StrokeTilingPattern!;
        double strokeHalf = primitive.StrokeWidth / 2D;
        double left = primitive.X - strokeHalf;
        double top = primitive.Y - strokeHalf;
        double width = primitive.Width + primitive.StrokeWidth;
        double height = primitive.Height + primitive.StrokeWidth;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
            left = Math.Min(primitive.X1, primitive.X2) - strokeHalf;
            top = Math.Min(primitive.Y1, primitive.Y2) - strokeHalf;
            width = Math.Abs(primitive.X2 - primitive.X1) + primitive.StrokeWidth;
            height = Math.Abs(primitive.Y2 - primitive.Y1) + primitive.StrokeWidth;
        }
        if (width <= 0D || height <= 0D) return;

        PdfPageClipPath strokeBounds = PdfPageClipPath.Rectangle(left, top, width, height);
        if (primitive.ClipPath.HasValue) {
            strokeBounds = PdfPageClipPath.ResolveActiveClip(primitive.ClipPath.Value, strokeBounds);
        }
        if (!TryFitClipToDrawing(strokeBounds, drawing.Width, drawing.Height, out PdfPageClipPath fitted)) return;

        OfficeDrawing tile = paint.Resource.Tile.Clone();
        if (paint.Tint.HasValue) tile.ApplyColorTint(paint.Tint.Value);
        var patternDrawing = new OfficeDrawing(fitted.Width, fitted.Height);
        OfficeTransform localTransform = paint.Transform.Then(OfficeTransform.Translate(-fitted.X, -fitted.Y));
        patternDrawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, fitted.Width, fitted.Height),
            paint.Resource.HorizontalStep,
            paint.Resource.VerticalStep,
            localTransform,
            maximumTileCount: 16384,
            opacity: paint.Opacity);

        OfficeDrawing strokeMask = CreatePatternStrokeMask(primitive, fitted);
        if (strokeMask.Elements.Count == 0) return;
        drawing.AddEffectDrawing(
            patternDrawing,
            OfficeTransform.Translate(fitted.X, fitted.Y),
            OfficeBlendMode.Normal,
            new OfficeDrawingSoftMask(strokeMask));
    }

    private static OfficeDrawing CreatePatternStrokeMask(PdfPageVisualPrimitive primitive, PdfPageClipPath fitted) {
        var rawMask = new OfficeDrawing(fitted.Width, fitted.Height);
        double sourceWidth = Math.Max(1D, primitive.Width);
        double sourceHeight = Math.Max(1D, primitive.Height);
        OfficeShape shape;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
            shape = OfficeShape.Rectangle(primitive.Width, primitive.Height);
        } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
            shape = OfficeShape.Line(
                primitive.X1 - primitive.X,
                primitive.Y1 - primitive.Y,
                primitive.X2 - primitive.X,
                primitive.Y2 - primitive.Y);
        } else {
            shape = OfficeShape.Path(primitive.PathCommands);
        }

        shape.FillColor = null;
        shape.StrokeColor = OfficeColor.White;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        var source = new OfficeDrawing(sourceWidth, sourceHeight);
        source.AddShape(shape, 0D, 0D);
        rawMask.AddDrawingForClippedRendering(source, primitive.X - fitted.X, primitive.Y - fitted.Y, null);

        if (fitted.IsRectangle) return rawMask;
        OfficeClipPath? activeClip = fitted.ToOfficeClipPath(fitted.X, fitted.Y);
        if (activeClip == null) return new OfficeDrawing(fitted.Width, fitted.Height);
        var clippedMask = new OfficeDrawing(fitted.Width, fitted.Height);
        clippedMask.AddClippedDrawing(rawMask, 0D, 0D, activeClip);
        return clippedMask;
    }

    private Dictionary<string, PdfPageTilingPatternResource> GetTilingPatternResources(
        PdfDictionary? resources,
        HashSet<string>? invokedPatternNames = null,
        TilingPatternResourceCache? resourceCache = null,
        TextContentParser.TextOutputBudget? textOutputBudget = null,
        PageContentBudget? pageContentBudget = null,
        Type3GlyphBudget? type3GlyphBudget = null,
        bool requireSupportedType3Content = false,
        int contentNestingDepth = 0,
        bool allowNestedPatternContent = false) {
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        type3GlyphBudget ??= new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        resourceCache ??= new TilingPatternResourceCache();
        var result = new Dictionary<string, PdfPageTilingPatternResource>(StringComparer.Ordinal);
        if (resources == null || !resources.Items.TryGetValue("Pattern", out PdfObject? patternObject)) return result;
        PdfDictionary? patterns = ResolveDictionary(patternObject);
        if (patterns == null) return result;
        foreach (KeyValuePair<string, PdfObject> entry in patterns.Items) {
            if (invokedPatternNames != null && !invokedPatternNames.Contains(entry.Key)) continue;
            if (ResolveObject(entry.Value) is not PdfStream stream) {
                continue;
            }

            var activeKey = (Stream: stream, Resources: resources);
            var cacheKey = (
                Stream: stream,
                Resources: resources,
                RequireSupportedType3Content: requireSupportedType3Content,
                AllowNestedPatternContent: allowNestedPatternContent || requireSupportedType3Content,
                ContentNestingDepth: contentNestingDepth);
            if (resourceCache.Resources.TryGetValue(cacheKey, out PdfPageTilingPatternResource? cached)) {
                if (cached != null) {
                    result[entry.Key] = cached;
                }
                continue;
            }

            if (!resourceCache.Active.Add(activeKey)) continue;
            PdfPageTilingPatternResource? pattern;
            try {
                int patternNestingDepth = contentNestingDepth + 1;
                EnsureContentNestingBudget(patternNestingDepth);
                pattern = TryReadTilingPattern(
                    stream,
                    resources,
                    textOutputBudget,
                    pageContentBudget,
                    type3GlyphBudget,
                    resourceCache,
                    requireSupportedType3Content,
                    allowNestedPatternContent,
                    patternNestingDepth,
                    out PdfPageTilingPatternResource? parsed)
                    ? parsed
                    : null;
                resourceCache.Resources[cacheKey] = pattern;
            } finally {
                resourceCache.Active.Remove(activeKey);
            }
            if (pattern != null) {
                result[entry.Key] = pattern;
            }
        }
        return result;
    }

    private bool TryReadTilingPattern(
        PdfObject? value,
        PdfDictionary parentResources,
        TextContentParser.TextOutputBudget? textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        TilingPatternResourceCache resourceCache,
        bool requireSupportedType3Content,
        bool allowNestedPatternContent,
        int contentNestingDepth,
        out PdfPageTilingPatternResource pattern) {
        pattern = null!;
        int? paintType;
        int? tilingType;
        if (ResolveObject(value) is not PdfStream stream ||
            TryReadInteger(stream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? typeObject) ? typeObject : null) != 1 ||
            ((paintType = TryReadInteger(stream.Dictionary.Items.TryGetValue("PaintType", out PdfObject? paintTypeObject) ? paintTypeObject : null)) != 1 && paintType != 2) ||
            ((tilingType = TryReadInteger(stream.Dictionary.Items.TryGetValue("TilingType", out PdfObject? tilingTypeObject) ? tilingTypeObject : null)) < 1 || tilingType > 3) ||
            !TryReadRectangle(stream.Dictionary.Items.TryGetValue("BBox", out PdfObject? boxObject) ? boxObject : null, out (double X1, double Y1, double X2, double Y2) box) ||
            ResolveObject(stream.Dictionary.Items.TryGetValue("XStep", out PdfObject? xStepObject) ? xStepObject : null) is not PdfNumber xStep ||
            ResolveObject(stream.Dictionary.Items.TryGetValue("YStep", out PdfObject? yStepObject) ? yStepObject : null) is not PdfNumber yStep ||
            !IsFinite(xStep.Value) || !IsFinite(yStep.Value) ||
            Math.Abs(xStep.Value) <= 0.0000001D || Math.Abs(yStep.Value) <= 0.0000001D) return false;
        double width = box.X2 - box.X1;
        double height = box.Y2 - box.Y1;
        if (width <= 0D || height <= 0D) return false;
        PdfDictionary? resources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? parentResources;
        int failureVersion = type3GlyphBudget.FailureVersion;
        bool allowNestedPatterns = (allowNestedPatternContent || requireSupportedType3Content) && paintType == 1;
        OfficeDrawing tile = CreatePatternTileDrawing(
            stream,
            resources,
            box,
            width,
            height,
            textOutputBudget ?? CreateTextOutputBudget(),
            pageContentBudget,
            type3GlyphBudget,
            resourceCache,
            requireSupportedType3Content,
            allowNestedPatterns,
            contentNestingDepth);
        if (type3GlyphBudget.FailureVersion != failureVersion) return false;
        Matrix2D matrix = stream.Dictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject)
            ? ReadPatternMatrix(matrixObject)
            : Matrix2D.Identity;
        if (!IsUsableTilingPatternMatrix(matrix)) return false;
        bool uncolored = paintType == 2;
        pattern = new PdfPageTilingPatternResource(tile, Math.Abs(xStep.Value), Math.Abs(yStep.Value), matrix, box.X1, box.Y2, uncolored);
        return true;
    }

    private static bool IsUsableTilingPatternMatrix(Matrix2D matrix) =>
        IsFinite(matrix.A) && IsFinite(matrix.B) && IsFinite(matrix.C) &&
        IsFinite(matrix.D) && IsFinite(matrix.E) && IsFinite(matrix.F) &&
        Math.Abs((matrix.A * matrix.D) - (matrix.B * matrix.C)) > 0.000000000001D;

    private OfficeDrawing CreatePatternTileDrawing(
        PdfStream stream,
        PdfDictionary? resources,
        (double X1, double Y1, double X2, double Y2) box,
        double width,
        double height,
        TextContentParser.TextOutputBudget textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        TilingPatternResourceCache resourceCache,
        bool requireSupportedType3Content,
        bool allowNestedPatterns,
        int contentNestingDepth) {
        var drawing = new OfficeDrawing(width, height);
        RegisterEmbeddedFonts(drawing, resources, new HashSet<PdfStream>(), 0);
        string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
        if (content.Length == 0) return drawing;
        Matrix2D transform = Matrix2D.Translation(-box.X1, -box.Y1);
        var activeForms = new HashSet<PdfStream>();
        var elements = new List<PdfPageDrawingElement>();
        var primitives = new List<PdfPageVisualPrimitive>();
        var renderedType3PaintOrders = new HashSet<double>();
        Type3SoftMaskValidationContext? softMaskValidation = requireSupportedType3Content
            ? type3GlyphBudget.GetOrCreateSoftMaskValidationContext(this)
            : null;
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            transform,
            width,
            height,
            primitives.Add,
            activeForms,
            renderedType3PaintOrders: renderedType3PaintOrders,
            type3GlyphBudget: type3GlyphBudget,
            contentNestingDepth: contentNestingDepth,
            includeTilingPatterns: allowNestedPatterns,
            requireSupportedType3Content: requireSupportedType3Content,
            allowSupportedType3Patterns: allowNestedPatterns,
            allowSupportedType3TransparencyGroups: requireSupportedType3Content,
            unrenderedPatternVisitor: requireSupportedType3Content || allowNestedPatterns
                ? null
                : _ => type3GlyphBudget.RecordFailure(),
            type3ImageVisitor: (placement, image, effect) => elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect)),
            type3PrimitiveVisitor: (primitive, effect) => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect)),
            type3GroupVisitor: (group, transform, paintOrder, key, effect) => elements.Add(PdfPageDrawingElement.FromGroup(group, transform, paintOrder, key, elements.Count).WithEffect(effect)),
            graphicsStateVisitor: softMaskValidation == null
                ? null
                : (state, stateTransform) => {
                    if (!CanDecodeType3SoftMask(
                            state.SoftMask,
                            stateTransform,
                            softMaskValidation.PageContentBudget,
                            softMaskValidation.ValidatedGroups,
                            softMaskValidation.Type3GlyphBudget,
                            contentNestingDepth + 1,
                            projectionPageWidth: width,
                            projectionPageHeight: height,
                            textOutputBudget: softMaskValidation.TextOutputBudget)) {
                        type3GlyphBudget.RecordFailure();
                    }
                },
            tilingPatternResourceCache: resourceCache,
            textOutputBudget: textOutputBudget,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root);
        for (int i = 0; i < primitives.Count; i++) elements.Add(PdfPageDrawingElement.FromPrimitive(primitives[i], elements.Count));

        var spans = new List<PdfTextSpan>();
        Dictionary<string, Func<byte[], int, string>> decoders = ResourceResolver.GetBudgetedFontDecodersForForm(stream.Dictionary, _objects);
        Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProviders(stream.Dictionary, _objects);
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        string transformedContent = WrapContentWithTransform(content, transform, out int transformedOffset);
        CollectTextAndForms(
            transformedContent,
            resources,
            decoders,
            widthProviders,
            fonts,
            spans,
            activeForms,
            height,
            paintOrderOffset: -transformedOffset,
            useLogicalTextFilters: false,
            contentNestingDepth: contentNestingDepth,
            textOutputBudget: textOutputBudget,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root,
            contentOrderOffset: -transformedOffset);
        for (int i = 0; i < spans.Count; i++) {
            if (renderedType3PaintOrders.Contains(spans[i].PaintOrder)) continue;
            elements.Add(PdfPageDrawingElement.FromText(spans[i], elements.Count));
        }

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(
            content,
            resources,
            0,
            transform,
            height,
            placements,
            activeForms,
            contentNestingDepth: contentNestingDepth,
            pageContentBudget: pageContentBudget);
        if (placements.Count > 0) {
            IReadOnlyList<PdfExtractedImage> images = GetImagesForResources(resources, 0, placements, colorizeImageMasks: true);
            for (int i = 0; i < placements.Count; i++) {
                PdfExtractedImage? image = FindImage(images, placements[i]);
                if (image != null) elements.Add(PdfPageDrawingElement.FromImage(placements[i], image, elements.Count));
            }
        }
        var enclosingEffects = new List<PdfPageDrawingEffectTransition>();
        CollectGraphicsEffectTransitions(
            content,
            resources,
            transform,
            height,
            enclosingEffects,
            new HashSet<PdfStream>(),
            PdfPageDrawingEffect.Default,
            contentNestingDepth: contentNestingDepth,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root);
        SortGraphicsEffectTransitions(enclosingEffects);
        OverlayDrawingEffects(elements, enclosingEffects);
        SortDrawingElements(elements);
        var softMasks = new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
        var activeSoftMasks = new HashSet<PdfStream>();
        for (int i = 0; i < elements.Count; i++) {
            AddDrawingElement(drawing, height, transform, elements[i], softMasks, activeSoftMasks, textOutputBudget, pageContentBudget, type3GlyphBudget);
        }
        return drawing;
    }

    private sealed class TilingPatternResourceCache {
        internal Dictionary<(PdfStream Stream, PdfDictionary Resources, bool RequireSupportedType3Content, bool AllowNestedPatternContent, int ContentNestingDepth), PdfPageTilingPatternResource?> Resources { get; } =
            new Dictionary<(PdfStream Stream, PdfDictionary Resources, bool RequireSupportedType3Content, bool AllowNestedPatternContent, int ContentNestingDepth), PdfPageTilingPatternResource?>();

        internal HashSet<(PdfStream Stream, PdfDictionary Resources)> Active { get; } =
            new HashSet<(PdfStream Stream, PdfDictionary Resources)>();
    }
}
