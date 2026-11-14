using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool IsSupportedOrdinaryTransparencyGroup(PdfDictionary form, PdfPageXObjectInvocation invocation) {
        if (ResolveEffectObject(form.Items.TryGetValue("Group", out var groupObject) ? groupObject : null) is not PdfDictionary group ||
            !HasOrdinaryTransparencyGroupDeclaration(group)) return false;
        if (group.Items.TryGetValue("K", out var knockout) && ResolveEffectObject(knockout) is not PdfNull and not PdfBoolean { Value: false }) return false;
        bool isolated = ResolveEffectObject(group.Items.TryGetValue("I", out var isolatedObject) ? isolatedObject : null) is PdfBoolean { Value: true };
        PdfObject? color = group.Items.TryGetValue("CS", out var colorObject) ? ResolveEffectObject(colorObject) : null;
        bool deviceRgb = color is null or PdfNull or PdfName { Name: "DeviceRGB" };
        if (isolated && deviceRgb) return TryReadFormMatrix(form, out _);
        // A group composited through a blend mode or constant alpha must be painted as one surface:
        // applying the invocation effect to each member is replaced by the members' own graphics state.
        // Members that paint only with Normal blending give the same result whether or not the group
        // is isolated, and an RGB ICC blending space is treated like DeviceRGB.
        if (invocation.BlendMode == OfficeBlendMode.Normal && (invocation.FillOpacity ?? 1D) >= 1D) return false;
        if (!deviceRgb && !IsThreeComponentIccColorSpace(color)) return false;
        if (!isolated && !PaintsOnlyWithNormalBlending(form)) return false;
        return TryReadFormMatrix(form, out _);
    }

    private bool IsThreeComponentIccColorSpace(PdfObject? color) =>
        color is PdfArray { Items.Count: 2 } array &&
        ResolveEffectObject(array.Items[0]) is PdfName { Name: "ICCBased" } &&
        ResolveEffectObject(array.Items[1]) is PdfStream profile &&
        ResolveEffectObject(profile.Dictionary.Items.TryGetValue("N", out PdfObject? components) ? components : null) is PdfNumber { Value: 3D };

    // Conservative: every graphics state the Form can select keeps Normal blending and no soft mask,
    // and the Form invokes no nested Forms whose own states would also apply.
    private bool PaintsOnlyWithNormalBlending(PdfDictionary form) {
        PdfDictionary? resources = ResolveDictionary(form.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null);
        if (resources == null) return true;
        if (ResolveDictionary(resources.Items.TryGetValue("ExtGState", out PdfObject? statesObject) ? statesObject : null) is PdfDictionary states) {
            foreach (PdfObject stateObject in states.Items.Values) {
                if (ResolveEffectObject(stateObject) is not PdfDictionary state) return false;
                if (state.Items.ContainsKey("BM") && ReadBlendMode(state) != OfficeBlendMode.Normal) return false;
                if (state.Items.TryGetValue("SMask", out PdfObject? mask) && ResolveEffectObject(mask) is not PdfName { Name: "None" }) return false;
            }
        }
        if (ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? xObjects) ? xObjects : null) is PdfDictionary objects) {
            foreach (PdfObject value in objects.Items.Values) {
                if (ResolveEffectObject(value) is PdfStream stream &&
                    ResolveEffectObject(stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null) is PdfName { Name: "Form" }) return false;
            }
        }
        return true;
    }

    // An isolated Form starts with opaque paint and no inherited blend/mask.
    // Its invocation opacity and effects apply to the completed surface once.
    private OfficeDrawing CreateOrdinaryTransparencyGroupDrawing(string content, PdfDictionary? resources,
        Matrix2D transform, double width, double height, PdfPageXObjectInvocation invocation,
        HashSet<PdfStream> activeForms, HashSet<PdfStream> activeType3Glyphs,
        RenderedType3TextTracker renderedType3PaintOrders, Type3GlyphBudget type3GlyphBudget,
        double paintOrderScale, TextContentParser.TextOutputBudget? textOutputBudget, PageContentBudget budget,
        PdfTextClippingBudget invocationClipping, PdfTextClippingBudget patternClipping,
        int nestingDepth, PdfContentOrderKey orderPrefix, PdfFontResourceSet fonts, PdfTextStateSnapshot textState) {
        var elements = new List<PdfPageDrawingElement>();
        CollectVisualPrimitivesAndForms(content, resources, transform, width, height,
            primitive => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count)),
            activeForms, activeType3Glyphs, renderedType3PaintOrders, type3GlyphBudget,
            invocation.PaintOrder, paintOrderScale * .000000001D,
            initialClipPath: invocation.ClipPath,
            initialFillColor: invocation.FillColor, initialFillColorSpace: invocation.FillColorSpace,
            initialFillPattern: invocation.FillPattern, initialFillPatternBaseColorSpace: invocation.FillPatternBaseColorSpace,
            initialFillOpacity: 1D,
            initialStrokeColor: invocation.StrokeColor, initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialStrokePattern: invocation.StrokePattern, initialStrokePatternBaseColorSpace: invocation.StrokePatternBaseColorSpace,
            initialStrokeOpacity: 1D, initialStrokeWidth: invocation.StrokeWidth,
            initialStrokeDashStyle: invocation.StrokeDashStyle, initialStrokeDashPattern: invocation.StrokeDashPattern,
            initialStrokeLineCap: invocation.StrokeLineCap, initialStrokeLineJoin: invocation.StrokeLineJoin,
            contentNestingDepth: nestingDepth + 1,
            type3ImageVisitor: (placement, image, effect) => elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect)),
            type3PrimitiveVisitor: (primitive, effect) => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect)),
            type3GroupVisitor: (group, mapping, order, key, effect) => elements.Add(PdfPageDrawingElement.FromGroup(group, mapping, order, key, elements.Count).WithEffect(effect)),
            textOutputBudget: textOutputBudget, pageContentBudget: budget,
            invocationTextClippingBudget: invocationClipping, patternTextClippingBudget: patternClipping,
            contentOrderPrefix: orderPrefix, initialRenderingIntent: invocation.RenderingIntent,
            initialFillColorSelection: invocation.FillColorSelection, initialStrokeColorSelection: invocation.StrokeColorSelection,
            initialTextState: textState, inheritedFontResources: fonts);

        var spans = new List<PdfTextSpan>();
        var actualTextScopes = new List<TextContentParser.DrawingActualTextScope>();
        string textContent = WrapContentWithTransform(content, transform, out int textOffset);
        CollectTextAndForms(textContent, resources, fonts.Decoders, fonts.WidthProviders, fonts.Fonts, spans, activeForms, height,
            invocation.PaintOrder, paintOrderScale * .000000001D, -textOffset,
            initialFillColor: invocation.FillColor, initialFillColorSpace: invocation.FillColorSpace,
            initialStrokeColor: invocation.StrokeColor, initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialFillOpacity: 1D, initialStrokeOpacity: 1D, initialClipPath: invocation.ClipPath,
            useLogicalTextFilters: false, includeArtifactText: true, contentNestingDepth: nestingDepth + 1,
            textOutputBudget: textOutputBudget, textClippingBudget: invocationClipping, pageContentBudget: budget,
            contentOrderPrefix: orderPrefix, contentOrderOffset: -textOffset,
            initialRenderingIntent: invocation.RenderingIntent, initialTextState: textState,
            initialTextRenderingMode: textState.TextRenderingMode, cancellationCheck: budget.CancellationToken.ThrowIfCancellationRequested,
            onCompletedDrawingActualText: actualTextScopes.Add);
        PdfPaintedGlyphRuns.SplitComplexRuns(spans, budget.ChargePositionedTextWorkCharacters, budget.CancellationToken);
        PdfArabicPaintedForms.Apply(spans, budget.CancellationToken);
        foreach (var span in spans) {
            if (!renderedType3PaintOrders.Contains(span.PaintOrder, span.ContentOrderKey))
                elements.Add(PdfPageDrawingElement.FromText(span.WithPageFontSize(), elements.Count));
        }

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(content, resources, 0, transform, height, placements, activeForms,
            invocation.FillColor, invocation.FillColorSpace, 1D, invocation.PaintOrder, paintOrderScale * .000000001D,
            initialClipPath: invocation.ClipPath, contentNestingDepth: nestingDepth + 1,
            pageContentBudget: budget, textClippingBudget: invocationClipping, contentOrderPrefix: orderPrefix,
            initialRenderingIntent: invocation.RenderingIntent);
        foreach (var placement in placements) {
            var image = GetImageForPlacement(resources, placement, colorizeImageMasks: true, budget);
            if (image != null) elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count));
        }
        var effects = new List<PdfPageDrawingEffectTransition>();
        CollectGraphicsEffectTransitions(content, resources, transform, height, effects, new HashSet<PdfStream>(),
            PdfPageDrawingEffect.Default, invocation.PaintOrder, paintOrderScale * .000000001D,
            initialClipPath: invocation.ClipPath, initialFillColor: invocation.FillColor, initialFillColorSpace: invocation.FillColorSpace,
            initialFillOpacity: 1D, initialStrokeColor: invocation.StrokeColor, initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialStrokeOpacity: 1D, initialStrokeWidth: invocation.StrokeWidth,
            initialStrokeDashStyle: invocation.StrokeDashStyle, initialStrokeDashPattern: invocation.StrokeDashPattern,
            initialStrokeLineCap: invocation.StrokeLineCap, initialStrokeLineJoin: invocation.StrokeLineJoin,
            initialRenderingIntent: invocation.RenderingIntent, contentNestingDepth: nestingDepth + 1,
            textClippingBudget: invocationClipping, pageContentBudget: budget, contentOrderPrefix: orderPrefix);
        SortGraphicsEffectTransitions(effects); OverlayDrawingEffects(elements, effects); SortDrawingElements(elements);

        var drawing = new OfficeDrawing(width, height);
        var registeredFonts = new Dictionary<(string Family, OfficeFontStyle Style), PdfFontResource>();
        foreach (PdfFontResource font in fonts.Fonts.Values) RegisterEmbeddedFont(drawing, font, registeredFonts);
        RegisterEmbeddedFonts(drawing, resources, new HashSet<PdfStream>(), nestingDepth + 1, registeredFonts, cancellationToken: budget.CancellationToken);
        ConfigureDrawingFonts(drawing, registeredFonts, budget);
        AddPaintedGlyphMappings(drawing, registeredFonts, elements, budget.PaintedGlyphMaps, budget, budget.CancellationToken);
        var masks = new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height, OfficeIccRenderingIntent Intent), OfficeDrawingSoftMask>();
        var activeMasks = budget.ActiveSoftMasks;
        var outputBudget = textOutputBudget ?? CreateTextOutputBudget();
        ReplayDrawingElements(drawing, elements, actualTextScopes, (destination, element) =>
            AddDrawingElement(destination, height, transform, element, masks, activeMasks,
                outputBudget, budget, type3GlyphBudget, invocationClipping, patternClipping, cancellationToken: budget.CancellationToken),
            budget.CancellationToken);
        ApplyPaintedGlyphMappings(drawing, registeredFonts, budget.PaintedGlyphMaps);
        double opacity = invocation.FillOpacity ?? 1D;
        if (opacity == 1D) return drawing;
        var surface = new OfficeDrawing(width, height);
        surface.AddEffectDrawing(drawing, OfficeTransform.Identity, opacity);
        return surface;
    }

    private PdfFontResourceSet CreateInheritedFormFontResources(PdfFontResourceSet parent, PdfDictionary? resources,
        PdfTextStateSnapshot inherited, out PdfTextStateSnapshot state, System.Threading.CancellationToken cancellationToken) {
        PdfFontResourceSet local = _fontResourceCache.GetOrCreate(resources, _objects, forDrawing: true, cancellationToken);
        var decoders = MergeDecoders(parent.Decoders, local.Decoders);
        var widths = MergeWidthProviders(parent.WidthProviders, local.WidthProviders);
        var fonts = MergeFonts(parent.Fonts, local.Fonts);
        state = PreserveInheritedFormFontState(inherited, parent.Decoders, parent.WidthProviders, parent.Fonts,
            local.Decoders, local.WidthProviders, local.Fonts, decoders, widths, fonts);
        return new PdfFontResourceSet(fonts, decoders, widths);
    }

    internal sealed partial class PageContentBudget {
        private HashSet<PdfContentOrderKey> _projectedTransparencyGroups = new();
        internal void RecordProjectedTransparencyGroup(PdfContentOrderKey key) => _projectedTransparencyGroups.Add(key);
        // Each walker tests the Form invocation before descending into its children.
        internal bool HasProjectedTransparencyGroup(PdfContentOrderKey? key) => key != null &&
            _projectedTransparencyGroups.Contains(key);

        internal IDisposable BeginTransparencyGroupProjectionScope() => new TransparencyGroupProjectionScope(this);

        private sealed class TransparencyGroupProjectionScope : IDisposable {
            private readonly PageContentBudget _owner;
            private readonly HashSet<PdfContentOrderKey> _previous;
            internal TransparencyGroupProjectionScope(PageContentBudget owner) {
                _owner = owner; _previous = owner._projectedTransparencyGroups;
                owner._projectedTransparencyGroups = new HashSet<PdfContentOrderKey>();
            }
            public void Dispose() => _owner._projectedTransparencyGroups = _previous;
        }
    }
}
