using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private OfficeBlendMode? ReadBlendMode(PdfDictionary state) {
        if (!state.Items.TryGetValue("BM", out PdfObject? value)) return null;
        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                OfficeBlendMode? candidate = MapBlendMode(ResolveObject(array.Items[index]) as PdfName);
                if (candidate.HasValue) return candidate;
            }
            return null;
        }
        return MapBlendMode(resolved as PdfName);
    }

    private static OfficeBlendMode? MapBlendMode(PdfName? name) {
        switch (name?.Name) {
            case "Normal": case "Compatible": return OfficeBlendMode.Normal;
            case "Multiply": return OfficeBlendMode.Multiply;
            case "Screen": return OfficeBlendMode.Screen;
            case "Overlay": return OfficeBlendMode.Overlay;
            case "Darken": return OfficeBlendMode.Darken;
            case "Lighten": return OfficeBlendMode.Lighten;
            case "ColorDodge": return OfficeBlendMode.ColorDodge;
            case "ColorBurn": return OfficeBlendMode.ColorBurn;
            case "HardLight": return OfficeBlendMode.HardLight;
            case "SoftLight": return OfficeBlendMode.SoftLight;
            case "Difference": return OfficeBlendMode.Difference;
            case "Exclusion": return OfficeBlendMode.Exclusion;
            case "Hue": return OfficeBlendMode.Hue;
            case "Saturation": return OfficeBlendMode.Saturation;
            case "Color": return OfficeBlendMode.Color;
            case "Luminosity": return OfficeBlendMode.Luminosity;
            default: return null;
        }
    }

    private PdfPageSoftMaskResource? ReadSoftMask(PdfDictionary state) {
        PdfObject? resolved = ResolveObject(state.Items.TryGetValue("SMask", out PdfObject? value) ? value : null);
        if (resolved is PdfName { Name: "None" } || resolved is not PdfDictionary mask ||
            ResolveObject(mask.Items.TryGetValue("G", out PdfObject? groupObject) ? groupObject : null) is not PdfStream group) return null;
        if (group.Dictionary.Get<PdfName>("Subtype")?.Name != "Form" ||
            ResolveDictionary(group.Dictionary.Items.TryGetValue("Group", out PdfObject? transparencyObject) ? transparencyObject : null)?.Get<PdfName>("S")?.Name != "Transparency" ||
            ResolveObject(mask.Items.TryGetValue("S", out PdfObject? modeObject) ? modeObject : null) is not PdfName modeName ||
            (modeName.Name != "Alpha" && modeName.Name != "Luminosity")) return null;
        OfficeSoftMaskMode mode = modeName.Name == "Luminosity" ? OfficeSoftMaskMode.Luminosity : OfficeSoftMaskMode.Alpha;
        OfficeColor backdrop = OfficeColor.Transparent;
        if (mode == OfficeSoftMaskMode.Luminosity &&
            ResolveObject(mask.Items.TryGetValue("BC", out PdfObject? backdropObject) ? backdropObject : null) is PdfArray components) {
            IReadOnlyList<double> values = ReadNumberArray(components);
            if (values.Count >= 3) backdrop = OfficeColor.FromRgb(ToColorByte(values[0]), ToColorByte(values[1]), ToColorByte(values[2]));
            else if (values.Count == 1) backdrop = OfficeColor.FromRgb(ToColorByte(values[0]), ToColorByte(values[0]), ToColorByte(values[0]));
        }
        return new PdfPageSoftMaskResource(group, mode, backdrop);
    }

    private IReadOnlyList<PdfPageDrawingEffectTransition> GetGraphicsEffectTransitions(Matrix2D pageTransform, double pageHeight) {
        var transitions = new List<PdfPageDrawingEffectTransition>();
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        string content = GetContentStreamContent();
        if (content.Length == 0) return Array.Empty<PdfPageDrawingEffectTransition>();
        var activeForms = new HashSet<PdfStream>();
        CollectGraphicsEffectTransitions(content, resources, pageTransform, pageHeight, transitions, activeForms, PdfPageDrawingEffect.Default);
        transitions.Sort(static (left, right) => left.PaintOrder.CompareTo(right.PaintOrder));
        return transitions.Count == 0 ? Array.Empty<PdfPageDrawingEffectTransition>() : transitions.AsReadOnly();
    }

    private void CollectGraphicsEffectTransitions(
        string content,
        PdfDictionary? resources,
        Matrix2D baseTransform,
        double pageHeight,
        List<PdfPageDrawingEffectTransition> transitions,
        HashSet<PdfStream> activeForms,
        PdfPageDrawingEffect initialEffect,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        double? initialFillOpacity = null,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialStrokeOpacity = null,
        double? initialStrokeWidth = null,
        OfficeStrokeDashStyle? initialStrokeDashStyle = null,
        OfficeStrokeLineCap? initialStrokeLineCap = null,
        OfficeStrokeLineJoin? initialStrokeLineJoin = null,
        int contentNestingDepth = 0) {
        EnsureContentNestingBudget(contentNestingDepth);
        Dictionary<string, PdfPageGraphicsStateResource> graphicsStates = GetGraphicsStateResources(resources);
        IReadOnlyList<PdfPageDrawingEffectTransition> local = PdfPageGraphicsEffectTimelineParser.Parse(
            content,
            graphicsStates,
            initialEffect,
            paintOrderBase,
            paintOrderScale,
            paintOrderOffset,
            _limits.MaxContentOperations,
            _limits.MaxContentNestingDepth,
            _limits.MaxContentOperands);
        transitions.AddRange(local);

        foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                     content,
                     baseTransform,
                     pageHeight,
                     graphicsStates,
                     GetColorSpaceResources(resources),
                     GetOptionalContentVisibility(resources),
                     initialFillColor,
                     initialFillColorSpace,
                     initialFillOpacity,
                     paintOrderBase,
                     paintOrderScale,
                     paintOrderOffset,
                     initialClipPath,
                     initialStrokeColor,
                     initialStrokeColorSpace,
                     initialStrokeOpacity,
                     initialStrokeWidth,
                     initialStrokeDashStyle,
                     initialStrokeLineCap,
                     initialStrokeLineJoin,
                     _limits.MaxContentOperations,
                     _limits.MaxContentNestingDepth,
                     _limits.MaxContentOperands)) {
            if (!TryGetFormStream(resources, invocation.Name, out PdfStream formStream) || !activeForms.Add(formStream)) continue;
            PdfPageDrawingEffect inherited = ResolveDrawingEffect(local, invocation.PaintOrder, initialEffect);
            try {
                PdfDictionary dictionary = formStream.Dictionary;
                PdfDictionary? formResources = ResolveDictionary(dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? resources;
                Matrix2D formTransform = ApplyFormMatrix(invocation.Transform, dictionary);
                string formContent = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(DecodeIfNeeded(formStream)), dictionary);
                CollectGraphicsEffectTransitions(
                    formContent,
                    formResources,
                    formTransform,
                    pageHeight,
                    transitions,
                    activeForms,
                    inherited,
                    invocation.PaintOrder,
                    paintOrderScale * 0.000000001D,
                    initialClipPath: invocation.ClipPath,
                    initialFillColor: invocation.FillColor,
                    initialFillColorSpace: invocation.FillColorSpace,
                    initialFillOpacity: invocation.FillOpacity,
                    initialStrokeColor: invocation.StrokeColor,
                    initialStrokeColorSpace: invocation.StrokeColorSpace,
                    initialStrokeOpacity: invocation.StrokeOpacity,
                    initialStrokeWidth: invocation.StrokeWidth,
                    initialStrokeDashStyle: invocation.StrokeDashStyle,
                    initialStrokeLineCap: invocation.StrokeLineCap,
                    initialStrokeLineJoin: invocation.StrokeLineJoin,
                    contentNestingDepth: contentNestingDepth + 1);
                transitions.Add(new PdfPageDrawingEffectTransition(invocation.PaintOrder + (Math.Abs(paintOrderScale) * 0.25D), inherited));
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private static PdfPageDrawingEffect ResolveDrawingEffect(
        IReadOnlyList<PdfPageDrawingEffectTransition> transitions,
        double paintOrder,
        PdfPageDrawingEffect? initial = null) {
        PdfPageDrawingEffect effect = initial ?? PdfPageDrawingEffect.Default;
        for (int i = 0; i < transitions.Count; i++) {
            if (transitions[i].PaintOrder > paintOrder) break;
            effect = transitions[i].Effect;
        }
        return effect;
    }

    private OfficeDrawingSoftMask GetOrCreateSoftMask(
        PdfPageSoftMaskResource resource,
        double width,
        double height,
        Matrix2D pageTransform,
        Dictionary<PdfPageSoftMaskResource, OfficeDrawingSoftMask> cache,
        TextContentParser.TextOutputBudget textOutputBudget) {
        if (cache.TryGetValue(resource, out OfficeDrawingSoftMask? existing)) return existing;
        OfficeDrawing drawing = CreateFormDrawing(resource.Group, width, height, pageTransform, textOutputBudget);
        var mask = new OfficeDrawingSoftMask(drawing, resource.Mode, backdropColor: resource.BackdropColor);
        cache[resource] = mask;
        return mask;
    }

    private OfficeDrawing CreateFormDrawing(
        PdfStream form,
        double width,
        double height,
        Matrix2D pageTransform,
        TextContentParser.TextOutputBudget textOutputBudget) {
        var drawing = new OfficeDrawing(width, height);
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        PdfDictionary? resources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? pageResources;
        RegisterEmbeddedFonts(drawing, resources, new HashSet<PdfStream>(), 0);
        string content = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(DecodeIfNeeded(form)), form.Dictionary);
        if (content.Length == 0) return drawing;
        Matrix2D transform = ApplyFormMatrix(pageTransform, form.Dictionary);
        var activeForms = new HashSet<PdfStream>();
        var elements = new List<PdfPageDrawingElement>();
        var primitives = new List<PdfPageVisualPrimitive>();
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            transform,
            width,
            height,
            primitives.Add,
            activeForms,
            textOutputBudget: textOutputBudget);
        for (int i = 0; i < primitives.Count; i++) elements.Add(PdfPageDrawingElement.FromPrimitive(primitives[i], elements.Count));

        var spans = new List<PdfTextSpan>();
        Dictionary<string, Func<byte[], string>> decoders = MergeDecoders(
            ResourceResolver.GetFontDecoders(_pageDict, _objects, _limits.MaxDecodedTextCharacters),
            ResourceResolver.GetFontDecodersForForm(form.Dictionary, _objects, _limits.MaxDecodedTextCharacters));
        Dictionary<string, Func<byte[], double>> widthProviders = MergeWidthProviders(ResourceResolver.GetFontWidthProviders(_pageDict, _objects), ResourceResolver.GetFontWidthProviders(form.Dictionary, _objects));
        Dictionary<string, PdfFontResource> fonts = MergeFonts(ResourceResolver.GetFontsForResources(pageResources, _objects), ResourceResolver.GetFontsForResources(resources, _objects));
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
            textOutputBudget: textOutputBudget);
        for (int i = 0; i < spans.Count; i++) elements.Add(PdfPageDrawingElement.FromText(spans[i], elements.Count));

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(content, resources, 0, transform, height, placements, activeForms);
        if (placements.Count > 0) {
            IReadOnlyList<PdfExtractedImage> images = GetImagesForResources(resources, 0, placements, colorizeImageMasks: true);
            for (int i = 0; i < placements.Count; i++) {
                PdfExtractedImage? image = FindImage(images, placements[i]);
                if (image != null) elements.Add(PdfPageDrawingElement.FromImage(placements[i], image, elements.Count));
            }
        }

        SortDrawingElements(elements);
        for (int i = 0; i < elements.Count; i++) AddDrawingElementCore(drawing, height, elements[i]);
        return drawing;
    }
}
