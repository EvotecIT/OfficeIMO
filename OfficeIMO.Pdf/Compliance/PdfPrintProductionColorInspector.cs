using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    internal static PdfPrintProductionColorEvidence Inspect(
        PdfReadDocument document,
        System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        int maximumObjectDepth = document.ReadOptions.Limits.MaxObjectNestingDepth;
        var contentStreams = new List<ContentStreamContext>();
        var imageDictionaries = new HashSet<PdfDictionary>();
        var imageContexts = new List<ImageContext>();
        var shadingContexts = new List<ShadingContext>();
        var graphicsStateDictionaries = new HashSet<PdfDictionary>();
        var shadingDictionaries = new HashSet<PdfDictionary>();
        int rgbImages = 0;
        int cmykImages = 0;
        int rgbShadings = 0;
        int cmykShadings = 0;
        int rgbTransparencyGroups = 0;
        int cmykTransparencyGroups = 0;
        int deviceIndependentColorUses = 0;
        int transparentImages = 0;
        int nonOpaqueStates = 0;
        int transparencyGroups = 0;
        int uninspectable = 0;

        int pageSequenceId = 0;
        foreach (PdfReadPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary dictionary = page.PageDictionary;

            ColorSpaceAliases pageAliases = ResolveColorSpaceAliases(
                dictionary,
                objects,
                inheritResources: true,
                maximumObjectDepth);
            PdfDictionary? pageResources = ResolveResourcesDictionary(
                dictionary,
                objects,
                inheritResources: true,
                maximumObjectDepth);
            int firstPageContext = contentStreams.Count;
            if (dictionary.Items.TryGetValue("Contents", out PdfObject? contents)) {
                if (!CollectStreams(
                    contents,
                    objects,
                    pageAliases,
                    pageResources,
                    contentStreams,
                    maximumObjectDepth,
                    pageSequenceId)) uninspectable++;
            }
            ReachableResourceCollection reachable = CollectReachableResourceContexts(
                contentStreams,
                firstPageContext,
                objects,
                imageContexts,
                shadingContexts,
                graphicsStateDictionaries,
                document.ReadOptions.Limits,
                cancellationToken);
            transparencyGroups += reachable.TransparencyGroupCount;
            if (!TryClassifyTransparencyGroup(
                    dictionary,
                    pageAliases,
                    objects,
                    maximumObjectDepth,
                    out bool isPageTransparencyGroup,
                    out ColorSpaceUsage? pageGroupUsage)) {
                uninspectable++;
            } else if (isPageTransparencyGroup) {
                transparencyGroups++;
                ApplyTransparencyGroupUsage(
                    pageGroupUsage,
                    ref rgbTransparencyGroups,
                    ref cmykTransparencyGroups,
                    ref deviceIndependentColorUses);
            }
            pageSequenceId++;
        }

        foreach (ContentStreamContext context in contentStreams) {
            if (!TryClassifyTransparencyGroup(
                    context.Stream.Dictionary,
                    context.Aliases,
                    objects,
                    maximumObjectDepth,
                    out bool isTransparencyGroup,
                    out ColorSpaceUsage? groupUsage)) {
                uninspectable++;
            } else if (isTransparencyGroup) {
                ApplyTransparencyGroupUsage(
                    groupUsage,
                    ref rgbTransparencyGroups,
                    ref cmykTransparencyGroups,
                    ref deviceIndependentColorUses);
            }
        }

        var softMaskOwners = new HashSet<PdfDictionary>();
        for (int contextIndex = 0; contextIndex < imageContexts.Count; contextIndex++) {
            ImageContext context = imageContexts[contextIndex];
            if (!context.Dictionary.Items.TryGetValue("SMask", out PdfObject? softMask) ||
                string.Equals(ResolveName(softMask, objects, maximumObjectDepth), "None", StringComparison.Ordinal)) {
                continue;
            }

            if (softMaskOwners.Add(context.Dictionary)) transparentImages++;
            if (ResolveObject(objects, softMask, 0, maximumObjectDepth) is not PdfStream softMaskStream ||
                !string.Equals(
                    ResolveName(
                        softMaskStream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null,
                        objects,
                        maximumObjectDepth),
                    "Image",
                    StringComparison.Ordinal)) {
                uninspectable++;
                continue;
            }

            AddImageContext(softMaskStream.Dictionary, context.Aliases, imageContexts);
        }

        foreach (ImageContext context in imageContexts) imageDictionaries.Add(context.Dictionary);
        foreach (ShadingContext context in shadingContexts) shadingDictionaries.Add(context.Dictionary);

        foreach (PdfDictionary image in imageDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (image.Items.TryGetValue("ImageMask", out PdfObject? imageMaskObject) &&
                ResolveObject(objects, imageMaskObject, 0, maximumObjectDepth) is PdfBoolean { Value: true }) {
                continue;
            }

            PdfObject? colorSpace = image.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                ? colorSpaceObject
                : null;
            bool hasContext = false;
            bool hasUnknownContext = false;
            bool usesRgb = false;
            bool usesCmyk = false;
            bool usesDeviceIndependentColor = false;
            foreach (ImageContext context in imageContexts) {
                if (!ReferenceEquals(context.Dictionary, image)) continue;
                hasContext = true;
                ColorSpaceUsage usage = ClassifyColorSpace(
                    colorSpace,
                    objects,
                    maximumObjectDepth,
                    context.Aliases);
                hasUnknownContext |= !usage.IsKnown;
                usesRgb |= usage.UsesDeviceRgb;
                usesCmyk |= usage.UsesDeviceCmyk;
                usesDeviceIndependentColor |= usage.UsesDeviceIndependent;
            }

            if (!hasContext) {
                ColorSpaceUsage usage = ClassifyColorSpace(colorSpace, objects, maximumObjectDepth);
                hasUnknownContext = !usage.IsKnown;
                usesRgb = usage.UsesDeviceRgb;
                usesCmyk = usage.UsesDeviceCmyk;
                usesDeviceIndependentColor = usage.UsesDeviceIndependent;
            }

            if (hasUnknownContext) uninspectable++;

            if (usesRgb) rgbImages++;
            if (usesCmyk) cmykImages++;
            if (usesDeviceIndependentColor) deviceIndependentColorUses++;
        }

        foreach (PdfDictionary shading in shadingDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? colorSpace = shading.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                ? colorSpaceObject
                : null;
            bool hasContext = false;
            bool hasUnknownContext = false;
            bool usesRgb = false;
            bool usesCmyk = false;
            bool usesDeviceIndependentColor = false;
            foreach (ShadingContext context in shadingContexts) {
                if (!ReferenceEquals(context.Dictionary, shading)) continue;
                hasContext = true;
                ColorSpaceUsage usage = ClassifyColorSpace(
                    colorSpace,
                    objects,
                    maximumObjectDepth,
                    context.Aliases);
                hasUnknownContext |= !usage.IsKnown;
                usesRgb |= usage.UsesDeviceRgb;
                usesCmyk |= usage.UsesDeviceCmyk;
                usesDeviceIndependentColor |= usage.UsesDeviceIndependent;
            }
            if (!hasContext) {
                ColorSpaceUsage usage = ClassifyColorSpace(colorSpace, objects, maximumObjectDepth);
                hasUnknownContext = !usage.IsKnown;
                usesRgb = usage.UsesDeviceRgb;
                usesCmyk = usage.UsesDeviceCmyk;
                usesDeviceIndependentColor = usage.UsesDeviceIndependent;
            }
            if (hasUnknownContext) uninspectable++;
            if (usesRgb) rgbShadings++;
            if (usesCmyk) cmykShadings++;
            if (usesDeviceIndependentColor) deviceIndependentColorUses++;
        }

        foreach (PdfDictionary graphicsState in graphicsStateDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryInspectGraphicsState(
                    graphicsState,
                    objects,
                    maximumObjectDepth,
                    out bool isNonOpaque)) {
                uninspectable++;
            } else if (isNonOpaque) {
                nonOpaqueStates++;
            }
        }

        int rgbOperators = 0;
        int cmykOperators = 0;
        var pageColorStates = new Dictionary<int, ContentColorState>();
        foreach (ContentStreamContext context in contentStreams) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfStream stream = context.Stream;
            ColorSpaceAliases aliases = context.Aliases;
            ContentColorState colorState;
            if (context.PageSequenceId is int currentPageSequenceId) {
                if (!pageColorStates.TryGetValue(currentPageSequenceId, out colorState!)) {
                    colorState = new ContentColorState();
                    pageColorStates[currentPageSequenceId] = colorState;
                }
            } else {
                colorState = new ContentColorState();
            }
            if (!StreamDecoder.TryDecode(
                    stream.Dictionary,
                    stream.Data,
                    document.ReadOptions.Limits.MaxDecodedStreamBytes,
                    out byte[] decoded,
                    objects)) {
                colorState.IsIncomplete = true;
                uninspectable++;
                continue;
            }

            try {
                bool contextWasUninspectable = colorState.IsIncomplete;
                PdfContentStreamInterpreter.Interpret(
                    PdfEncoding.Latin1GetString(decoded),
                    document.ReadOptions.Limits.MaxContentOperations,
                    operation => {
                        if (operation.HasInvalidOperands) {
                            contextWasUninspectable = true;
                            return;
                        }

                        switch (operation.Name) {
                            case "q":
                                colorState.Stack.Push(new ContentColorStateSnapshot(
                                    colorState.FillUsesDeviceRgb,
                                    colorState.StrokeUsesDeviceRgb,
                                    colorState.FillUsesDeviceCmyk,
                                    colorState.StrokeUsesDeviceCmyk,
                                    colorState.FillUsesDeviceIndependentColor,
                                    colorState.StrokeUsesDeviceIndependentColor));
                                break;
                            case "Q":
                                if (colorState.Stack.Count > 0) {
                                    ContentColorStateSnapshot restored = colorState.Stack.Pop();
                                    colorState.FillUsesDeviceRgb = restored.FillUsesDeviceRgb;
                                    colorState.StrokeUsesDeviceRgb = restored.StrokeUsesDeviceRgb;
                                    colorState.FillUsesDeviceCmyk = restored.FillUsesDeviceCmyk;
                                    colorState.StrokeUsesDeviceCmyk = restored.StrokeUsesDeviceCmyk;
                                    colorState.FillUsesDeviceIndependentColor = restored.FillUsesDeviceIndependentColor;
                                    colorState.StrokeUsesDeviceIndependentColor = restored.StrokeUsesDeviceIndependentColor;
                                } else {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "rg":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultRgb ?? ColorSpaceUsage.DeviceRgb,
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "RG":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultRgb ?? ColorSpaceUsage.DeviceRgb,
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "k":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultCmyk ?? ColorSpaceUsage.DeviceCmyk,
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "K":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultCmyk ?? ColorSpaceUsage.DeviceCmyk,
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "g":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultGray ?? ColorSpaceUsage.DeviceGray,
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "G":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultGray ?? ColorSpaceUsage.DeviceGray,
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "cs":
                                ApplyColorSpaceUsage(
                                    ClassifySelectedColorSpace(operation, aliases),
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "CS":
                                ApplyColorSpaceUsage(
                                    ClassifySelectedColorSpace(operation, aliases),
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "sc":
                            case "scn":
                                if (colorState.FillUsesDeviceRgb) rgbOperators++;
                                if (colorState.FillUsesDeviceCmyk) cmykOperators++;
                                if (colorState.FillUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "SC":
                            case "SCN":
                                if (colorState.StrokeUsesDeviceRgb) rgbOperators++;
                                if (colorState.StrokeUsesDeviceCmyk) cmykOperators++;
                                if (colorState.StrokeUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                        }

                        if (operation.InlineImage != null) {
                            PdfDictionary inlineImage = operation.InlineImage.Dictionary;
                            if (inlineImage.Items.TryGetValue("ColorSpace", out PdfObject? inlineColorSpace)) {
                                ColorSpaceUsage usage = ClassifyColorSpace(
                                    inlineColorSpace,
                                    objects,
                                    maximumObjectDepth,
                                    aliases,
                                    normalizeInlineImageAbbreviations: true);
                                if (!usage.IsKnown) contextWasUninspectable = true;
                                if (usage.UsesDeviceRgb) rgbImages++;
                                if (usage.UsesDeviceCmyk) cmykImages++;
                                if (usage.UsesDeviceIndependent) deviceIndependentColorUses++;
                            } else {
                                bool isImageMask = inlineImage.Items.TryGetValue("ImageMask", out PdfObject? imageMaskObject) &&
                                    ResolveObject(objects, imageMaskObject, 0, maximumObjectDepth) is PdfBoolean { Value: true };
                                if (!isImageMask) contextWasUninspectable = true;
                            }
                        }
                    },
                    maxNestingDepth: document.ReadOptions.Limits.MaxContentNestingDepth,
                    maxOperands: document.ReadOptions.Limits.MaxContentOperands,
                    dispatchInvalidOperations: true);
                if (contextWasUninspectable || context.ResourceInspectionIncomplete) {
                    colorState.IsIncomplete = true;
                    uninspectable++;
                }
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is PdfReadLimitException ||
                exception is FormatException) {
                colorState.IsIncomplete = true;
                uninspectable++;
            }
        }

        return new PdfPrintProductionColorEvidence(
            rgbOperators,
            cmykOperators,
            rgbImages,
            cmykImages,
            rgbShadings,
            cmykShadings,
            rgbTransparencyGroups,
            cmykTransparencyGroups,
            deviceIndependentColorUses,
            transparentImages,
            nonOpaqueStates,
            transparencyGroups,
            uninspectable);
    }

    private static bool CollectStreams(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        ColorSpaceAliases aliases,
        PdfDictionary? resources,
        List<ContentStreamContext> streams,
        int maximumObjectDepth,
        int pageSequenceId) {
        bool complete = true;
        var pending = new Stack<(PdfObject Value, int Depth)>();
        var inspectedArrays = new HashSet<PdfArray>();
        pending.Push((value, 0));
        while (pending.Count > 0) {
            (PdfObject candidate, int depth) = pending.Pop();
            ThrowIfObjectDepthExceeded(depth, maximumObjectDepth);
            PdfObject? resolved = ResolveObject(
                objects,
                candidate,
                depth,
                maximumObjectDepth,
                out int resolvedDepth);
            if (resolved is PdfStream stream) {
                AddContentStream(stream, aliases, resources, streams, pageSequenceId: pageSequenceId);
            } else if (resolved is PdfArray array && inspectedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            } else {
                complete = false;
            }
        }
        return complete;
    }

    private static ColorSpaceAliases ResolveColorSpaceAliases(
        PdfDictionary owner,
        Dictionary<int, PdfIndirectObject> objects,
        bool inheritResources,
        int maximumObjectDepth) {
        PdfDictionary? resources = ResolveResourcesDictionary(
            owner,
            objects,
            inheritResources,
            maximumObjectDepth);
        return resources == null
            ? new ColorSpaceAliases()
            : CreateColorSpaceAliases(resources, objects, maximumObjectDepth);
    }

    private static PdfDictionary? ResolveResourcesDictionary(
        PdfDictionary owner,
        Dictionary<int, PdfIndirectObject> objects,
        bool inheritResources,
        int maximumObjectDepth) {
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? current = owner;
        int currentDepth = 0;
        while (current != null && visited.Add(current)) {
            if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                ResolveObject(
                    objects,
                    resourcesObject,
                    currentDepth + 1,
                    maximumObjectDepth) is PdfDictionary resources) {
                return resources;
            }

            if (!inheritResources ||
                !current.Items.TryGetValue("Parent", out PdfObject? parentObject) ||
                ResolveObject(
                    objects,
                    parentObject,
                    currentDepth + 1,
                    maximumObjectDepth,
                    out int parentDepth) is not PdfDictionary parent) {
                break;
            }

            current = parent;
            currentDepth = parentDepth;
        }

        return null;
    }

    private static ColorSpaceAliases CreateColorSpaceAliases(
        PdfDictionary resources,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        var aliases = new ColorSpaceAliases();
        CollectResourceColorSpaces(
            resources,
            objects,
            aliases.Rgb,
            aliases.Cmyk,
            aliases.Gray,
            aliases.Pattern,
            aliases.DeviceIndependent,
            maximumObjectDepth);
        if (resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) &&
            ResolveObject(objects, colorSpacesObject, 0, maximumObjectDepth) is PdfDictionary colorSpaces) {
            aliases.DefaultRgb = ResolveDefaultColorSpaceUsage(colorSpaces, "DefaultRGB", objects, maximumObjectDepth);
            aliases.DefaultCmyk = ResolveDefaultColorSpaceUsage(colorSpaces, "DefaultCMYK", objects, maximumObjectDepth);
            aliases.DefaultGray = ResolveDefaultColorSpaceUsage(colorSpaces, "DefaultGray", objects, maximumObjectDepth);
        }
        return aliases;
    }

    private static ColorSpaceUsage? ResolveDefaultColorSpaceUsage(
        PdfDictionary colorSpaces,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (!colorSpaces.Items.TryGetValue(key, out PdfObject? value)) return null;
        PdfObject? resolved = ResolveObject(objects, value, 0, maximumObjectDepth);
        if (resolved is PdfArray array &&
            array.Items.Count > 0 &&
            ResolveObject(objects, array.Items[0], 0, maximumObjectDepth) is PdfName family &&
            string.Equals(family.Name, "ICCBased", StringComparison.Ordinal)) {
            int expectedComponents = string.Equals(key, "DefaultRGB", StringComparison.Ordinal) ? 3 :
                string.Equals(key, "DefaultCMYK", StringComparison.Ordinal) ? 4 : 1;
            if (array.Items.Count != 2 ||
                ResolveObject(objects, array.Items[1], 0, maximumObjectDepth) is not PdfStream profile ||
                !TryResolveNumber(profile.Dictionary, "N", objects, maximumObjectDepth, out double components) ||
                components != expectedComponents) {
                return ColorSpaceUsage.Unknown;
            }
        }
        return ClassifyColorSpace(value, objects, maximumObjectDepth);
    }

    private static ColorSpaceUsage ClassifyColorSpace(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        ColorSpaceAliases? aliases = null,
        bool normalizeInlineImageAbbreviations = false) {
        ColorSpaceUsage usage = ClassifyColorSpaceCore(
            value,
            objects,
            maximumObjectDepth,
            aliases,
            normalizeInlineImageAbbreviations,
            new HashSet<PdfArray>(),
            0);
        if (aliases == null) return usage;

        if (usage.UsesDeviceRgb && aliases.DefaultRgb != null) usage = usage.ReplaceDeviceRgb(aliases.DefaultRgb);
        if (usage.UsesDeviceCmyk && aliases.DefaultCmyk != null) usage = usage.ReplaceDeviceCmyk(aliases.DefaultCmyk);
        if (usage.UsesDeviceGray && aliases.DefaultGray != null) usage = usage.ReplaceDeviceGray(aliases.DefaultGray);
        return usage;
    }

    private static ColorSpaceUsage ClassifyColorSpaceCore(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        ColorSpaceAliases? aliases,
        bool normalizeInlineImageAbbreviations,
        HashSet<PdfArray> activeArrays,
        int depth) {
        if (value == null) return ColorSpaceUsage.Unknown;
        PdfObject? resolved = ResolveObject(objects, value, depth, maximumObjectDepth, out int resolvedDepth);
        if (resolved is PdfName name) {
            string colorSpaceName = normalizeInlineImageAbbreviations
                ? NormalizeInlineImageColorSpaceName(name.Name)
                : name.Name;
            if (string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal) ||
                aliases?.Rgb.Contains(colorSpaceName) == true) return ColorSpaceUsage.DeviceRgb;
            if (string.Equals(colorSpaceName, "DeviceCMYK", StringComparison.Ordinal) ||
                aliases?.Cmyk.Contains(colorSpaceName) == true) return ColorSpaceUsage.DeviceCmyk;
            if (string.Equals(colorSpaceName, "DeviceGray", StringComparison.Ordinal) ||
                aliases?.Gray.Contains(colorSpaceName) == true) return ColorSpaceUsage.DeviceGray;
            if (string.Equals(colorSpaceName, "Pattern", StringComparison.Ordinal) ||
                aliases?.Pattern.Contains(colorSpaceName) == true) return ColorSpaceUsage.Pattern;
            if (aliases?.DeviceIndependent.Contains(colorSpaceName) == true) {
                return ColorSpaceUsage.DeviceIndependent;
            }
            return ColorSpaceUsage.Unknown;
        }
        if (resolved is not PdfArray array || !activeArrays.Add(array)) return ColorSpaceUsage.Unknown;

        try {
            if (array.Items.Count < 1 ||
                ResolveObject(objects, array.Items[0], resolvedDepth + 1, maximumObjectDepth) is not PdfName family) {
                return ColorSpaceUsage.Unknown;
            }

            switch (family.Name) {
                case "CalGray":
                case "CalRGB":
                case "Lab":
                    return array.Items.Count == 2 &&
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is PdfDictionary
                        ? ColorSpaceUsage.DeviceIndependent
                        : ColorSpaceUsage.Unknown;
                case "ICCBased":
                    if (array.Items.Count != 2 ||
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is not PdfStream profile ||
                        !TryResolveNumber(profile.Dictionary, "N", objects, maximumObjectDepth, out double components) ||
                        (components != 1D && components != 3D && components != 4D)) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ColorSpaceUsage.DeviceIndependent;
                case "Indexed":
                case "I":
                    if (array.Items.Count != 4 ||
                        !TryResolveBoundedInteger(array.Items[2], objects, maximumObjectDepth, 0, 255) ||
                        ResolveObject(objects, array.Items[3], resolvedDepth + 1, maximumObjectDepth) is not (PdfStringObj or PdfStream)) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ClassifyColorSpaceCore(
                        array.Items[1], objects, maximumObjectDepth, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1);
                case "Separation":
                    if (array.Items.Count != 4 ||
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is not PdfName ||
                        !IsColorSpaceFunction(array.Items[3], objects, maximumObjectDepth, resolvedDepth + 1)) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ClassifyColorSpaceCore(
                        array.Items[2], objects, maximumObjectDepth, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1);
                case "DeviceN":
                    if ((array.Items.Count != 4 && array.Items.Count != 5) ||
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is not PdfArray colorants ||
                        colorants.Items.Count < 1 ||
                        !AllArrayItemsAreNames(colorants, objects, maximumObjectDepth, resolvedDepth + 1) ||
                        !IsColorSpaceFunction(array.Items[3], objects, maximumObjectDepth, resolvedDepth + 1) ||
                        (array.Items.Count == 5 &&
                            ResolveObject(objects, array.Items[4], resolvedDepth + 1, maximumObjectDepth) is not PdfDictionary)) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ClassifyColorSpaceCore(
                        array.Items[2], objects, maximumObjectDepth, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1);
                case "Pattern":
                    if (array.Items.Count == 1) return ColorSpaceUsage.Pattern;
                    if (array.Items.Count != 2) return ColorSpaceUsage.Unknown;
                    ColorSpaceUsage baseUsage = ClassifyColorSpaceCore(
                        array.Items[1], objects, maximumObjectDepth, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1);
                    return baseUsage.IsKnown ? baseUsage.WithPattern() : ColorSpaceUsage.Unknown;
                default:
                    return ColorSpaceUsage.Unknown;
            }
        } finally {
            activeArrays.Remove(array);
        }
    }

    private static bool TryResolveBoundedInteger(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int minimum,
        int maximum) {
        if (ResolveObject(objects, value, 0, maximumObjectDepth) is not PdfNumber number ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value) ||
            number.Value != Math.Truncate(number.Value)) return false;
        return number.Value >= minimum && number.Value <= maximum;
    }

    private static bool IsColorSpaceFunction(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int depth) =>
        ResolveObject(objects, value, depth, maximumObjectDepth) is PdfDictionary or PdfStream;

    private static bool AllArrayItemsAreNames(
        PdfArray values,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int depth) {
        for (int index = 0; index < values.Items.Count; index++) {
            if (ResolveObject(objects, values.Items[index], depth, maximumObjectDepth) is not PdfName) return false;
        }
        return true;
    }

    private static void ApplyColorSpaceUsage(
        ColorSpaceUsage usage,
        ref bool usesRgb,
        ref bool usesCmyk,
        ref bool usesDeviceIndependent,
        ref int rgbOperators,
        ref int cmykOperators,
        ref int deviceIndependentColorUses,
        ref bool contextWasUninspectable) {
        usesRgb = usage.UsesDeviceRgb;
        usesCmyk = usage.UsesDeviceCmyk;
        usesDeviceIndependent = usage.UsesDeviceIndependent;
        if (!usage.IsKnown) contextWasUninspectable = true;
        if (usesRgb) rgbOperators++;
        if (usesCmyk) cmykOperators++;
        if (usesDeviceIndependent) deviceIndependentColorUses++;
    }

    private static ColorSpaceUsage ClassifySelectedColorSpace(
        PdfContentOperation operation,
        ColorSpaceAliases aliases) {
        if (operation.Operands.Count != 1 || operation.Operands[0] is not string colorSpaceName) {
            return ColorSpaceUsage.Unknown;
        }

        bool usesRgb = string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal) ||
            aliases.Rgb.Contains(colorSpaceName);
        bool usesCmyk = string.Equals(colorSpaceName, "DeviceCMYK", StringComparison.Ordinal) ||
            aliases.Cmyk.Contains(colorSpaceName);
        bool usesGray = string.Equals(colorSpaceName, "DeviceGray", StringComparison.Ordinal) ||
            aliases.Gray.Contains(colorSpaceName);
        bool usesPattern = string.Equals(colorSpaceName, "Pattern", StringComparison.Ordinal) ||
            aliases.Pattern.Contains(colorSpaceName);
        // CalGray, CalRGB, Lab, and ICCBased are array color-space families, not
        // directly selectable built-in names. They are valid here only through a
        // resource alias whose array definition was inspected above.
        bool usesDeviceIndependent = aliases.DeviceIndependent.Contains(colorSpaceName);
        ColorSpaceUsage usage = new(
            usesRgb || usesCmyk || usesGray || usesPattern || usesDeviceIndependent,
            usesRgb,
            usesCmyk,
            usesGray,
            usesPattern,
            usesDeviceIndependent);
        if (usesRgb && aliases.DefaultRgb != null) usage = usage.ReplaceDeviceRgb(aliases.DefaultRgb);
        if (usesCmyk && aliases.DefaultCmyk != null) usage = usage.ReplaceDeviceCmyk(aliases.DefaultCmyk);
        if (usesGray && aliases.DefaultGray != null) usage = usage.ReplaceDeviceGray(aliases.DefaultGray);
        return usage;
    }

    private static void CollectResourceColorSpaces(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<string> rgbAliases,
        HashSet<string> cmykAliases,
        HashSet<string> grayAliases,
        HashSet<string> patternAliases,
        HashSet<string> deviceIndependentAliases,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            ResolveObject(objects, colorSpacesObject, 0, maximumObjectDepth) is not PdfDictionary colorSpaces) return;

        var aliases = new ColorSpaceAliases();
        CopyAliases(rgbAliases, cmykAliases, grayAliases, patternAliases, deviceIndependentAliases, aliases);
        bool changed;
        do {
            changed = false;
            foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
                ColorSpaceUsage usage = ClassifyColorSpace(entry.Value, objects, maximumObjectDepth, aliases);
                if (!usage.IsKnown) continue;
                changed |= AddAliasUsage(entry.Key, usage, aliases);
            }
        } while (changed);

        CopyAliases(aliases.Rgb, aliases.Cmyk, aliases.Gray, aliases.Pattern, aliases.DeviceIndependent,
            rgbAliases, cmykAliases, grayAliases, patternAliases, deviceIndependentAliases);
    }

    private static string NormalizeInlineImageColorSpaceName(string name) => name switch {
        "RGB" => "DeviceRGB",
        "CMYK" => "DeviceCMYK",
        "G" => "DeviceGray",
        _ => name
    };

    private static void CopyAliases(
        HashSet<string> rgb,
        HashSet<string> cmyk,
        HashSet<string> gray,
        HashSet<string> pattern,
        HashSet<string> deviceIndependent,
        ColorSpaceAliases target) {
        target.Rgb.UnionWith(rgb);
        target.Cmyk.UnionWith(cmyk);
        target.Gray.UnionWith(gray);
        target.Pattern.UnionWith(pattern);
        target.DeviceIndependent.UnionWith(deviceIndependent);
    }

    private static void CopyAliases(
        HashSet<string> rgb,
        HashSet<string> cmyk,
        HashSet<string> gray,
        HashSet<string> pattern,
        HashSet<string> deviceIndependent,
        HashSet<string> targetRgb,
        HashSet<string> targetCmyk,
        HashSet<string> targetGray,
        HashSet<string> targetPattern,
        HashSet<string> targetDeviceIndependent) {
        targetRgb.UnionWith(rgb);
        targetCmyk.UnionWith(cmyk);
        targetGray.UnionWith(gray);
        targetPattern.UnionWith(pattern);
        targetDeviceIndependent.UnionWith(deviceIndependent);
    }

    private static bool AddAliasUsage(string name, ColorSpaceUsage usage, ColorSpaceAliases aliases) {
        bool changed = false;
        if (usage.UsesDeviceRgb) changed |= aliases.Rgb.Add(name);
        if (usage.UsesDeviceCmyk) changed |= aliases.Cmyk.Add(name);
        if (usage.UsesDeviceGray) changed |= aliases.Gray.Add(name);
        if (usage.UsesPattern) changed |= aliases.Pattern.Add(name);
        if (usage.UsesDeviceIndependent) changed |= aliases.DeviceIndependent.Add(name);
        return changed;
    }

    private static string? ResolveName(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        value == null ? null : (ResolveObject(objects, value, 0, maximumObjectDepth) as PdfName)?.Name;

    private static bool TryInspectGraphicsState(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out bool isNonOpaque) {
        isNonOpaque = false;
        if (!TryInspectAlpha(dictionary, "ca", objects, maximumObjectDepth, ref isNonOpaque) ||
            !TryInspectAlpha(dictionary, "CA", objects, maximumObjectDepth, ref isNonOpaque)) return false;
        if (dictionary.Items.TryGetValue("BM", out PdfObject? blendObject)) {
            if (!TryInspectBlendMode(blendObject, objects, maximumObjectDepth, out bool isNonNormal)) return false;
            if (isNonNormal) isNonOpaque = true;
        }
        if (dictionary.Items.TryGetValue("SMask", out PdfObject? softMask) &&
            !string.Equals(ResolveName(softMask, objects, maximumObjectDepth), "None", StringComparison.Ordinal)) {
            isNonOpaque = true;
        }
        return true;
    }

    private static bool TryInspectAlpha(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        ref bool isNonOpaque) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate)) return true;
        if (ResolveObject(objects, candidate, 0, maximumObjectDepth) is not PdfNumber number ||
            double.IsNaN(number.Value) || double.IsInfinity(number.Value) ||
            number.Value < 0D || number.Value > 1D) return false;
        if (number.Value != 1D) isNonOpaque = true;
        return true;
    }

    private static bool TryClassifyTransparencyGroup(
        PdfDictionary dictionary,
        ColorSpaceAliases aliases,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out bool isTransparencyGroup,
        out ColorSpaceUsage? usage) {
        isTransparencyGroup = false;
        usage = null;
        if (!dictionary.Items.TryGetValue("Group", out PdfObject? groupObject)) return true;
        if (ResolveObject(objects, groupObject, 0, maximumObjectDepth) is not PdfDictionary group ||
            ResolveName(group.Items.TryGetValue("S", out PdfObject? subtype) ? subtype : null, objects, maximumObjectDepth) is not string subtypeName) {
            return false;
        }
        if (!string.Equals(subtypeName, "Transparency", StringComparison.Ordinal)) return false;
        isTransparencyGroup = true;
        if (!group.Items.TryGetValue("CS", out PdfObject? colorSpace)) return true;
        usage = ClassifyColorSpace(colorSpace, objects, maximumObjectDepth, aliases);
        return usage.IsKnown;
    }

    private static void ApplyTransparencyGroupUsage(
        ColorSpaceUsage? usage,
        ref int rgbTransparencyGroups,
        ref int cmykTransparencyGroups,
        ref int deviceIndependentColorUses) {
        if (usage == null) return;
        if (usage.UsesDeviceRgb) rgbTransparencyGroups++;
        if (usage.UsesDeviceCmyk) cmykTransparencyGroups++;
        if (usage.UsesDeviceIndependent) deviceIndependentColorUses++;
    }

    private static bool TryInspectBlendMode(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out bool isNonNormal) {
        isNonNormal = false;
        var pending = new Stack<(PdfObject Value, int Depth)>();
        var inspectedArrays = new HashSet<PdfArray>();
        pending.Push((value, 0));
        while (pending.Count > 0) {
            (PdfObject candidate, int depth) = pending.Pop();
            ThrowIfObjectDepthExceeded(depth, maximumObjectDepth);
            PdfObject? resolved = ResolveObject(
                objects,
                candidate,
                depth,
                maximumObjectDepth,
                out int resolvedDepth);
            if (resolved is PdfName name) {
                if (!TryClassifyBlendModeName(name.Name, out isNonNormal)) continue;
                return true;
            } else if (resolved is PdfArray array) {
                if (array.Items.Count == 0) return false;
                if (!inspectedArrays.Add(array)) continue;
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            } else {
                continue;
            }
        }
        return false;
    }

    private static bool TryClassifyBlendModeName(string name, out bool isNonNormal) {
        isNonNormal = false;
        if (string.Equals(name, "Normal", StringComparison.Ordinal) ||
            string.Equals(name, "Compatible", StringComparison.Ordinal)) return true;
        switch (name) {
            case "Multiply":
            case "Screen":
            case "Overlay":
            case "Darken":
            case "Lighten":
            case "ColorDodge":
            case "ColorBurn":
            case "HardLight":
            case "SoftLight":
            case "Difference":
            case "Exclusion":
            case "Hue":
            case "Saturation":
            case "Color":
            case "Luminosity":
                isNonNormal = true;
                return true;
            default:
                return false;
        }
    }

    private static void ThrowIfObjectDepthExceeded(int depth, int maximumObjectDepth) {
        if (depth > maximumObjectDepth) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.ObjectNestingDepth,
                maximumObjectDepth,
                depth);
        }
    }

    private static PdfObject? ResolveObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        int depth,
        int maximumObjectDepth) =>
        ResolveObject(objects, value, depth, maximumObjectDepth, out _);

    private static PdfObject? ResolveObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        int depth,
        int maximumObjectDepth,
        out int resolvedDepth) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        PdfObject? resolved = value;
        resolvedDepth = depth;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) return null;
            resolvedDepth++;
            ThrowIfObjectDepthExceeded(resolvedDepth, maximumObjectDepth);
            resolved = indirect.Value;
        }
        return resolved;
    }

    private static bool TryResolveNumber(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out double value) {
        value = 0D;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
            ResolveObject(objects, candidate, 0, maximumObjectDepth) is not PdfNumber number) {
            return false;
        }

        value = number.Value;
        return true;
    }

    private sealed class ColorSpaceAliases {
        internal HashSet<string> Rgb { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Cmyk { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Gray { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Pattern { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> DeviceIndependent { get; } = new(StringComparer.Ordinal);
        internal ColorSpaceUsage? DefaultRgb { get; set; }
        internal ColorSpaceUsage? DefaultCmyk { get; set; }
        internal ColorSpaceUsage? DefaultGray { get; set; }

        internal bool SetEquals(ColorSpaceAliases other) =>
            Rgb.SetEquals(other.Rgb) &&
            Cmyk.SetEquals(other.Cmyk) &&
            Gray.SetEquals(other.Gray) &&
            Pattern.SetEquals(other.Pattern) &&
            DeviceIndependent.SetEquals(other.DeviceIndependent) &&
            Equals(DefaultRgb, other.DefaultRgb) &&
            Equals(DefaultCmyk, other.DefaultCmyk) &&
            Equals(DefaultGray, other.DefaultGray);
    }

    private sealed record ColorSpaceUsage(
        bool IsKnown,
        bool UsesDeviceRgb,
        bool UsesDeviceCmyk,
        bool UsesDeviceGray,
        bool UsesPattern,
        bool UsesDeviceIndependent) {
        internal static ColorSpaceUsage DeviceRgb { get; } = new(true, true, false, false, false, false);
        internal static ColorSpaceUsage DeviceCmyk { get; } = new(true, false, true, false, false, false);
        internal static ColorSpaceUsage DeviceGray { get; } = new(true, false, false, true, false, false);
        internal static ColorSpaceUsage Pattern { get; } = new(true, false, false, false, true, false);
        internal static ColorSpaceUsage DeviceIndependent { get; } = new(true, false, false, false, false, true);
        internal static ColorSpaceUsage Unknown { get; } = new(false, false, false, false, false, false);

        internal ColorSpaceUsage ReplaceDeviceRgb(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                replacement.UsesDeviceRgb,
                UsesDeviceCmyk || replacement.UsesDeviceCmyk,
                UsesDeviceGray || replacement.UsesDeviceGray,
                UsesPattern || replacement.UsesPattern,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent);

        internal ColorSpaceUsage ReplaceDeviceCmyk(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                UsesDeviceRgb || replacement.UsesDeviceRgb,
                replacement.UsesDeviceCmyk,
                UsesDeviceGray || replacement.UsesDeviceGray,
                UsesPattern || replacement.UsesPattern,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent);

        internal ColorSpaceUsage ReplaceDeviceGray(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                UsesDeviceRgb || replacement.UsesDeviceRgb,
                UsesDeviceCmyk || replacement.UsesDeviceCmyk,
                replacement.UsesDeviceGray,
                UsesPattern || replacement.UsesPattern,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent);

        internal ColorSpaceUsage WithPattern() =>
            new(IsKnown, UsesDeviceRgb, UsesDeviceCmyk, UsesDeviceGray, true, UsesDeviceIndependent);
    }

    private sealed class ContentStreamContext {
        internal ContentStreamContext(
            PdfStream stream,
            ColorSpaceAliases aliases,
            PdfDictionary? resources,
            PdfObject? inheritedFontObject,
            int? pageSequenceId) {
            Stream = stream;
            Aliases = aliases;
            Resources = resources;
            InheritedFontObject = inheritedFontObject;
            PageSequenceId = pageSequenceId;
        }

        internal PdfStream Stream { get; }
        internal ColorSpaceAliases Aliases { get; }
        internal PdfDictionary? Resources { get; }
        internal PdfObject? InheritedFontObject { get; }
        internal int? PageSequenceId { get; }
        internal bool ResourceInspectionIncomplete { get; set; }
    }

    private sealed class ContentColorState {
        internal bool FillUsesDeviceRgb;
        internal bool StrokeUsesDeviceRgb;
        internal bool FillUsesDeviceCmyk;
        internal bool StrokeUsesDeviceCmyk;
        internal bool FillUsesDeviceIndependentColor;
        internal bool StrokeUsesDeviceIndependentColor;
        internal bool IsIncomplete;
        internal Stack<ContentColorStateSnapshot> Stack { get; } = new();
    }

    private readonly record struct ContentColorStateSnapshot(
        bool FillUsesDeviceRgb,
        bool StrokeUsesDeviceRgb,
        bool FillUsesDeviceCmyk,
        bool StrokeUsesDeviceCmyk,
        bool FillUsesDeviceIndependentColor,
        bool StrokeUsesDeviceIndependentColor);

    private sealed record ReachableResourceCollection(int TransparencyGroupCount);

    private sealed record ImageContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);

    private sealed record ShadingContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);
}
