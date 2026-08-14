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
        if (!IsWithinTilingPatternLimit(paint, fitted)) return;
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
        double strokePadding = GetTilingPatternStrokePadding(primitive);
        double left = primitive.X - strokePadding;
        double top = primitive.Y - strokePadding;
        double width = primitive.Width + (strokePadding * 2D);
        double height = primitive.Height + (strokePadding * 2D);
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
            left = Math.Min(primitive.X1, primitive.X2) - strokePadding;
            top = Math.Min(primitive.Y1, primitive.Y2) - strokePadding;
            width = Math.Abs(primitive.X2 - primitive.X1) + (strokePadding * 2D);
            height = Math.Abs(primitive.Y2 - primitive.Y1) + (strokePadding * 2D);
        }
        if (width <= 0D || height <= 0D) return;

        PdfPageClipPath strokeBounds = PdfPageClipPath.Rectangle(left, top, width, height);
        if (primitive.ClipPath.HasValue) {
            strokeBounds = PdfPageClipPath.ResolveActiveClip(primitive.ClipPath.Value, strokeBounds);
        }
        if (!TryFitClipToDrawing(strokeBounds, drawing.Width, drawing.Height, out PdfPageClipPath fitted)) return;
        if (!IsWithinTilingPatternLimit(paint, fitted)) return;

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

    private static bool CanRenderTilingPatterns(PdfPageVisualPrimitive primitive, double drawingWidth, double drawingHeight) {
        if (primitive.FillTilingPattern is PdfPageTilingPatternPaint fillPaint &&
            (IsMagnifyingTilingPatternTransform(fillPaint.Transform) ||
             (TryGetTilingPatternFillBounds(primitive, drawingWidth, drawingHeight, out PdfPageClipPath fillBounds) &&
              (!fillBounds.IsExact || !IsWithinTilingPatternLimit(fillPaint, fillBounds))))) return false;
        if (primitive.StrokeTilingPattern is PdfPageTilingPatternPaint strokePaint && primitive.StrokeWidth > 0D &&
            (IsMagnifyingTilingPatternTransform(strokePaint.Transform) ||
             (TryGetTilingPatternStrokeBounds(primitive, drawingWidth, drawingHeight, out PdfPageClipPath strokeBounds) &&
              (!strokeBounds.IsExact || !IsWithinTilingPatternLimit(strokePaint, strokeBounds))))) return false;
        return true;
    }

    private static bool IsMagnifyingTilingPatternTransform(OfficeTransform transform) {
        // The largest singular value is the maximum scale applied in any direction.
        double firstLengthSquared = (transform.M11 * transform.M11) + (transform.M12 * transform.M12);
        double secondLengthSquared = (transform.M21 * transform.M21) + (transform.M22 * transform.M22);
        double dot = (transform.M11 * transform.M21) + (transform.M12 * transform.M22);
        double trace = firstLengthSquared + secondLengthSquared;
        double discriminant = ((firstLengthSquared - secondLengthSquared) * (firstLengthSquared - secondLengthSquared)) +
            (4D * dot * dot);
        if (!IsFinite(trace) || !IsFinite(discriminant)) return true;
        double largestScaleSquared = (trace + Math.Sqrt(Math.Max(0D, discriminant))) / 2D;
        return !IsFinite(largestScaleSquared) || largestScaleSquared > 1.000000000001D;
    }

    private static bool TryGetTilingPatternFillBounds(
        PdfPageVisualPrimitive primitive,
        double drawingWidth,
        double drawingHeight,
        out PdfPageClipPath fitted) {
        fitted = default;
        if (primitive.Width <= 0D || primitive.Height <= 0D) return false;
        PdfPageClipPath shapeClip;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
            shapeClip = PdfPageClipPath.Rectangle(primitive.X, primitive.Y, primitive.Width, primitive.Height);
        } else if (!PdfPageClipPath.TryCreatePath(primitive.PathCommands, primitive.FillRule, out shapeClip)) {
            return false;
        }
        if (primitive.ClipPath.HasValue) shapeClip = PdfPageClipPath.ResolveActiveClip(primitive.ClipPath.Value, shapeClip);
        return TryFitClipToDrawing(shapeClip, drawingWidth, drawingHeight, out fitted);
    }

    internal static bool TryGetTilingPatternStrokeBounds(
        PdfPageVisualPrimitive primitive,
        double drawingWidth,
        double drawingHeight,
        out PdfPageClipPath fitted) {
        fitted = default;
        double strokePadding = GetTilingPatternStrokePadding(primitive);
        double left = primitive.X - strokePadding;
        double top = primitive.Y - strokePadding;
        double width = primitive.Width + (strokePadding * 2D);
        double height = primitive.Height + (strokePadding * 2D);
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
            left = Math.Min(primitive.X1, primitive.X2) - strokePadding;
            top = Math.Min(primitive.Y1, primitive.Y2) - strokePadding;
            width = Math.Abs(primitive.X2 - primitive.X1) + (strokePadding * 2D);
            height = Math.Abs(primitive.Y2 - primitive.Y1) + (strokePadding * 2D);
        }
        if (width <= 0D || height <= 0D) return false;
        PdfPageClipPath bounds = PdfPageClipPath.Rectangle(left, top, width, height);
        if (primitive.ClipPath.HasValue) bounds = PdfPageClipPath.ResolveActiveClip(primitive.ClipPath.Value, bounds);
        return TryFitClipToDrawing(bounds, drawingWidth, drawingHeight, out fitted);
    }

    private static double GetTilingPatternStrokePadding(PdfPageVisualPrimitive primitive) {
        double strokeHalf = primitive.StrokeWidth / 2D;
        double padding = strokeHalf;
        OfficeStrokeLineJoin join = primitive.StrokeLineJoin ?? OfficeStrokeLineJoin.Miter;
        if (join == OfficeStrokeLineJoin.Miter) {
            padding = strokeHalf * 10D;
        } else if (join == OfficeStrokeLineJoin.Bevel) {
            padding = strokeHalf * Math.Sqrt(2D);
        }
        if (primitive.StrokeLineCap == OfficeStrokeLineCap.Square) {
            padding = Math.Max(padding, strokeHalf * Math.Sqrt(2D));
        }
        return padding;
    }

    private static bool IsWithinTilingPatternLimit(PdfPageTilingPatternPaint paint, PdfPageClipPath fitted) {
        OfficeTransform transform = paint.Transform.Then(OfficeTransform.Translate(-fitted.X, -fitted.Y));
        if (!transform.TryInvert(out OfficeTransform inverse)) return false;
        OfficePoint topLeft = inverse.TransformPoint(new OfficePoint(0D, 0D));
        OfficePoint topRight = inverse.TransformPoint(new OfficePoint(fitted.Width, 0D));
        OfficePoint bottomRight = inverse.TransformPoint(new OfficePoint(fitted.Width, fitted.Height));
        OfficePoint bottomLeft = inverse.TransformPoint(new OfficePoint(0D, fitted.Height));
        double minX = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
        double maxX = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
        double minY = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
        double maxY = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
        if (!TryGetTileRange(minX, maxX, 0D, paint.Resource.Tile.Width, paint.Resource.HorizontalStep, out long firstColumn, out long lastColumn) ||
            !TryGetTileRange(minY, maxY, 0D, paint.Resource.Tile.Height, paint.Resource.VerticalStep, out long firstRow, out long lastRow)) return false;
        double columns = (double)lastColumn - firstColumn + 1D;
        double rows = (double)lastRow - firstRow + 1D;
        return columns <= 0D || rows <= 0D ||
            IsFinite(columns) && IsFinite(rows) &&
            columns <= MaximumPatternVisibilityTiles && rows <= MaximumPatternVisibilityTiles &&
            columns * rows <= MaximumPatternVisibilityTiles;
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
            requireSupportedType3Content &&
                ResolveEffectObject(stream.Dictionary.Items.TryGetValue("Type", out PdfObject? patternTypeObject) ? patternTypeObject : null) is not PdfName { Name: "Pattern" } ||
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
        if (requireSupportedType3Content && !HasExactType3PatternBox(boxObject)) return false;
        PdfDictionary? resources;
        if (!TryResolveStrictResources(stream.Dictionary, parentResources, out resources)) return false;
        int failureVersion = type3GlyphBudget.FailureVersion;
        bool uncolored = paintType == 2;
        bool allowNestedPatterns = (allowNestedPatternContent || requireSupportedType3Content) && paintType == 1;
        bool rejectImageContent = requireSupportedType3Content && paintType == 2;
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
            rejectImageContent,
            allowNestedPatterns,
            uncolored,
            contentNestingDepth,
            out bool consumesInheritedLineState,
            out bool hasMalformedStrictInvocation);
        if (type3GlyphBudget.FailureVersion != failureVersion) return false;
        Matrix2D matrix;
        if (stream.Dictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject)) {
            if (requireSupportedType3Content) {
                if (!TryReadStrictPatternMatrix(matrixObject, out matrix)) return false;
            } else {
                matrix = ReadPatternMatrix(matrixObject);
            }
        } else {
            matrix = Matrix2D.Identity;
        }
        if (!IsUsableTilingPatternMatrix(matrix)) return false;
        pattern = new PdfPageTilingPatternResource(tile, Math.Abs(xStep.Value), Math.Abs(yStep.Value), matrix, box.X1, box.Y2, uncolored, consumesInheritedLineState, hasMalformedStrictInvocation);
        return true;
    }

    private bool HasExactType3PatternBox(PdfObject? value) {
        PdfArray? array = ResolveArray(value);
        if (array == null || array.Items.Count != 4) return false;
        for (int index = 0; index < array.Items.Count; index++) {
            if (ResolveObject(array.Items[index]) is not PdfNumber number || !IsFinite(number.Value)) return false;
        }
        return true;
    }

    private static bool IsUsableTilingPatternMatrix(Matrix2D matrix) {
        if (!IsFinite(matrix.A) || !IsFinite(matrix.B) || !IsFinite(matrix.C) ||
            !IsFinite(matrix.D) || !IsFinite(matrix.E) || !IsFinite(matrix.F)) return false;
        double determinant = (matrix.A * matrix.D) - (matrix.B * matrix.C);
        if (!IsFinite(determinant) || Math.Abs(determinant) <= 0.000000000001D) return false;
        double inverseA = matrix.D / determinant;
        double inverseB = -matrix.B / determinant;
        double inverseC = -matrix.C / determinant;
        double inverseD = matrix.A / determinant;
        double inverseE = -((inverseA * matrix.E) + (inverseC * matrix.F));
        double inverseF = -((inverseB * matrix.E) + (inverseD * matrix.F));
        return IsFinite(inverseA) && IsFinite(inverseB) && IsFinite(inverseC) &&
            IsFinite(inverseD) && IsFinite(inverseE) && IsFinite(inverseF);
    }

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
        bool rejectImageContent,
        bool allowNestedPatterns,
        bool rejectColorOperators,
        int contentNestingDepth,
        out bool consumesInheritedLineState,
        out bool hasMalformedStrictInvocation) {
        var drawing = new OfficeDrawing(width, height);
        consumesInheritedLineState = false;
        hasMalformedStrictInvocation = false;
        RegisterEmbeddedFonts(drawing, resources, new HashSet<PdfStream>(), 0);
        string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
        if (content.Length == 0) return drawing;
        hasMalformedStrictInvocation = HasMalformedStrictInvocation(
            content,
            resources,
            pageContentBudget,
            new HashSet<PdfStream>(),
            contentNestingDepth,
            rejectColorOperators);
        consumesInheritedLineState = ConsumesInheritedPatternLineState(
            content,
            resources,
            pageContentBudget,
            new HashSet<PdfStream>(),
            contentNestingDepth);
        if (requireSupportedType3Content && (consumesInheritedLineState || hasMalformedStrictInvocation)) {
            type3GlyphBudget.RecordFailure();
            return drawing;
        }
        Matrix2D transform = Matrix2D.Translation(-box.X1, -box.Y1);
        var activeForms = new HashSet<PdfStream>();
        var elements = new List<PdfPageDrawingElement>();
        var primitives = new List<PdfPageVisualPrimitive>();
        var renderedType3PaintOrders = new RenderedType3TextTracker();
        Type3SoftMaskValidationContext? softMaskValidation = requireSupportedType3Content
            ? type3GlyphBudget.GetOrCreateSoftMaskValidationContext(this)
            : null;
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            transform,
            width,
            height,
            primitive => {
                if (requireSupportedType3Content && !CanRenderTilingPatterns(primitive, width, height)) {
                    type3GlyphBudget.RecordFailure();
                } else {
                    primitives.Add(primitive);
                }
            },
            activeForms,
            renderedType3PaintOrders: renderedType3PaintOrders,
            type3GlyphBudget: type3GlyphBudget,
            contentNestingDepth: contentNestingDepth,
            includeTilingPatterns: allowNestedPatterns,
            requireSupportedType3Content: requireSupportedType3Content,
            allowSupportedType3Patterns: allowNestedPatterns,
            allowSupportedType3TransparencyGroups: requireSupportedType3Content,
            requireNestedType3Uncolored: rejectImageContent,
            unrenderedPatternVisitor: requireSupportedType3Content || allowNestedPatterns
                ? null
                : _ => type3GlyphBudget.RecordFailure(),
            type3ImageVisitor: (placement, image, effect) => {
                if (requireSupportedType3Content &&
                    !TryCreateImageProjection(
                        placement,
                        height,
                        width,
                        height,
                        out _,
                        allowAxisAlignedFallback: false)) {
                    type3GlyphBudget.RecordFailure();
                    return;
                }
                elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect));
            },
            type3PrimitiveVisitor: (primitive, effect) => {
                if (requireSupportedType3Content && !CanRenderTilingPatterns(primitive, width, height)) {
                    type3GlyphBudget.RecordFailure();
                } else {
                    elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect));
                }
            },
            type3GroupVisitor: (group, transform, paintOrder, key, effect) => elements.Add(PdfPageDrawingElement.FromGroup(group, transform, paintOrder, key, elements.Count).WithEffect(effect)),
            graphicsStateVisitor: softMaskValidation == null
                ? null
                : (state, stateTransform, fillColor, strokeColor, hasFillPattern, hasStrokePattern, stateNestingDepth) => {
                    if (!CanDecodeType3SoftMask(
                            state.SoftMask,
                            stateTransform,
                            softMaskValidation.PageContentBudget,
                            softMaskValidation.ValidatedGroups,
                            softMaskValidation.Type3GlyphBudget,
                            stateNestingDepth + 1,
                            projectionPageWidth: width,
                            projectionPageHeight: height,
                            textOutputBudget: softMaskValidation.TextOutputBudget,
                            inheritedFillColor: fillColor,
                            inheritedStrokeColor: strokeColor,
                            hasInheritedFillPattern: hasFillPattern,
                            hasInheritedStrokePattern: hasStrokePattern,
                            inheritedGraphicsState: state)) {
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
            if (renderedType3PaintOrders.Contains(spans[i].PaintOrder, spans[i].ContentOrderKey)) continue;
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
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root);
        if (placements.Count > 0) {
            for (int i = 0; i < placements.Count; i++) {
                PdfExtractedImage? image = GetImageForPlacement(resources, placements[i], colorizeImageMasks: true);
                if (requireSupportedType3Content &&
                    (rejectImageContent || image == null || !IsSupportedType3Image(placements[i], image!, resources) || image!.HasUnresolvedTransparencyMask ||
                     !TryCreateImageProjection(placements[i], height, width, height, out _, allowAxisAlignedFallback: false))) {
                    type3GlyphBudget.RecordFailure();
                    continue;
                }
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

    private bool ConsumesInheritedPatternLineState(
        string content,
        PdfDictionary? resources,
        PageContentBudget pageContentBudget,
        HashSet<PdfStream> activeForms,
        int contentNestingDepth,
        PatternLineState authoredState = default,
        Dictionary<(PdfStream Stream, PdfDictionary? Resources, int LineStateMask), bool>? type3LineStateCache = null) {
        type3LineStateCache ??= new Dictionary<(PdfStream Stream, PdfDictionary? Resources, int LineStateMask), bool>();
        var stateStack = new Stack<PatternLineState>();
        bool consumesInheritedState = false;
        Dictionary<string, PdfPageGraphicsStateResource> graphicsStates = GetGraphicsStateResources(resources);
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        PdfContentStreamInterpreter.Interpret(
            content,
            _limits.MaxContentOperations,
            operation => {
                if (consumesInheritedState) return;
                switch (operation.Name) {
                    case "q":
                        stateStack.Push(authoredState);
                        break;
                    case "Q":
                        authoredState = stateStack.Count > 0 ? stateStack.Pop() : default;
                        break;
                    case "w" when operation.Operands.Count == 1:
                        authoredState.Width = true;
                        break;
                    case "d" when operation.Operands.Count == 2:
                        authoredState.Dash = true;
                        break;
                    case "J" when operation.Operands.Count == 1:
                        authoredState.Cap = true;
                        break;
                    case "j" when operation.Operands.Count == 1:
                        authoredState.Join = true;
                        break;
                    case "Tf" when operation.Operands.Count == 2 && operation.Operands[0] is string fontName:
                        authoredState.FontName = fontName;
                        break;
                    case "gs" when operation.Operands.Count == 1 && operation.Operands[0] is string name && graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource graphicsState):
                        authoredState.Width |= graphicsState.StrokeWidth.HasValue;
                        authoredState.Dash |= graphicsState.StrokeDashStyle.HasValue;
                        authoredState.Cap |= graphicsState.StrokeLineCap.HasValue;
                        authoredState.Join |= graphicsState.StrokeLineJoin.HasValue;
                        break;
                    case "S": case "s": case "B": case "B*": case "b": case "b*":
                        consumesInheritedState = !authoredState.IsComplete;
                        break;
                    case "Do" when operation.Operands.Count == 1 && operation.Operands[0] is string xObjectName:
                        if (TryResolvePatternForm(resources, xObjectName, out PdfStream form) && activeForms.Add(form)) {
                            try {
                                EnsureContentNestingBudget(contentNestingDepth + 1);
                                PdfDictionary? formResources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourcesObject) ? formResourcesObject : null) ?? resources;
                                string formContent = PdfEncoding.Latin1GetString(pageContentBudget.Decode(form));
                                consumesInheritedState = ConsumesInheritedPatternLineState(
                                    formContent,
                                    formResources,
                                    pageContentBudget,
                                    activeForms,
                                    contentNestingDepth + 1,
                                    authoredState,
                                    type3LineStateCache);
                            } finally {
                                activeForms.Remove(form);
                            }
                        }
                        break;
                    case "Tj": case "TJ": case "'": case "\"":
                        consumesInheritedState = PatternType3TextConsumesInheritedLineState(
                            operation,
                            fonts,
                            pageContentBudget,
                            activeForms,
                            contentNestingDepth,
                            authoredState,
                            type3LineStateCache);
                        break;
                }
            },
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands);
        return consumesInheritedState;
    }

    private bool PatternType3TextConsumesInheritedLineState(
        PdfContentOperation operation,
        Dictionary<string, PdfFontResource> fonts,
        PageContentBudget pageContentBudget,
        HashSet<PdfStream> activeStreams,
        int contentNestingDepth,
        PatternLineState authoredState,
        Dictionary<(PdfStream Stream, PdfDictionary? Resources, int LineStateMask), bool> type3LineStateCache) {
        if (authoredState.IsComplete ||
            authoredState.FontName is not string fontName ||
            !fonts.TryGetValue(fontName, out PdfFontResource? font) ||
            font.Type3 is not PdfType3FontResource type3) return false;

        foreach (byte[] bytes in GetShownTextBytes(operation)) {
            for (int index = 0; index < bytes.Length; index++) {
                if (!type3.TryGetGlyph(bytes[index], out PdfStream glyph)) continue;
                var cacheKey = (glyph, type3.Resources, authoredState.LineStateMask);
                if (type3LineStateCache.TryGetValue(cacheKey, out bool cached)) {
                    if (cached) return true;
                    continue;
                }
                if (!activeStreams.Add(glyph)) continue;
                try {
                    EnsureContentNestingBudget(contentNestingDepth + 1);
                    string glyphContent = PdfEncoding.Latin1GetString(pageContentBudget.Decode(glyph));
                    bool consumes = ConsumesInheritedPatternLineState(
                        glyphContent,
                        type3.Resources,
                        pageContentBudget,
                        activeStreams,
                        contentNestingDepth + 1,
                        authoredState,
                        type3LineStateCache);
                    type3LineStateCache[cacheKey] = consumes;
                    if (consumes) return true;
                } finally {
                    activeStreams.Remove(glyph);
                }
            }
        }
        return false;
    }

    private static IEnumerable<byte[]> GetShownTextBytes(PdfContentOperation operation) {
        if (string.Equals(operation.Name, "TJ", StringComparison.Ordinal)) {
            if (operation.Operands.Count == 1 && operation.Operands[0] is List<object> items) {
                for (int index = 0; index < items.Count; index++) {
                    if (items[index] is byte[] bytes) yield return bytes;
                }
            }
            yield break;
        }
        if (operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is byte[] text) {
            yield return text;
        }
    }

    private bool TryResolvePatternForm(PdfDictionary? resources, string name, out PdfStream form) {
        form = null!;
        if (resources == null ||
            ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) ? xObjectsObject : null) is not PdfDictionary xObjects ||
            !xObjects.Items.TryGetValue(name, out PdfObject? formObject) ||
            ResolveObject(formObject) is not PdfStream candidate ||
            !string.Equals(candidate.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            return false;
        }
        form = candidate;
        return true;
    }

    private struct PatternLineState {
        internal bool Width;
        internal bool Dash;
        internal bool Cap;
        internal bool Join;
        internal string? FontName;
        internal bool IsComplete => Width && Dash && Cap && Join;
        internal int LineStateMask =>
            (Width ? 1 : 0) | (Dash ? 2 : 0) | (Cap ? 4 : 0) | (Join ? 8 : 0);
    }

    private sealed class TilingPatternResourceCache {
        internal Dictionary<(PdfStream Stream, PdfDictionary Resources, bool RequireSupportedType3Content, bool AllowNestedPatternContent, int ContentNestingDepth), PdfPageTilingPatternResource?> Resources { get; } =
            new Dictionary<(PdfStream Stream, PdfDictionary Resources, bool RequireSupportedType3Content, bool AllowNestedPatternContent, int ContentNestingDepth), PdfPageTilingPatternResource?>();

        internal HashSet<(PdfStream Stream, PdfDictionary Resources)> Active { get; } =
            new HashSet<(PdfStream Stream, PdfDictionary Resources)>();
    }
}
