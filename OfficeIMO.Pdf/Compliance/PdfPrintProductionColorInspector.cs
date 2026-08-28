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
        int maximumDecodedStreamBytes = document.ReadOptions.Limits.MaxDecodedStreamBytes;
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
                maximumObjectDepth,
                maximumDecodedStreamBytes);
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
                    maximumDecodedStreamBytes,
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
                    maximumDecodedStreamBytes,
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
                    maximumDecodedStreamBytes,
                    context.Aliases);
                hasUnknownContext |= !usage.IsKnown;
                usesRgb |= usage.UsesDeviceRgb;
                usesCmyk |= usage.UsesDeviceCmyk;
                usesDeviceIndependentColor |= usage.UsesDeviceIndependent;
            }

            if (!hasContext) {
                ColorSpaceUsage usage = ClassifyColorSpace(colorSpace, objects, maximumObjectDepth, maximumDecodedStreamBytes);
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
                    maximumDecodedStreamBytes,
                    context.Aliases);
                hasUnknownContext |= !usage.IsKnown ||
                    !IsStructurallyInspectableShading(
                        context,
                        usage.ComponentCount,
                        objects,
                        maximumObjectDepth,
                        maximumDecodedStreamBytes);
                usesRgb |= usage.UsesDeviceRgb;
                usesCmyk |= usage.UsesDeviceCmyk;
                usesDeviceIndependentColor |= usage.UsesDeviceIndependent;
            }
            if (!hasContext) {
                ColorSpaceUsage usage = ClassifyColorSpace(colorSpace, objects, maximumObjectDepth, maximumDecodedStreamBytes);
                hasUnknownContext = !usage.IsKnown ||
                    !IsStructurallyInspectableShading(
                        new ShadingContext(shading, null, new ColorSpaceAliases()),
                        usage.ComponentCount,
                        objects,
                        maximumObjectDepth,
                        maximumDecodedStreamBytes);
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
        for (int contextIndex = 0; contextIndex < contentStreams.Count; contextIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            ContentStreamContext context = contentStreams[contextIndex];
            ColorSpaceAliases aliases = context.Aliases;
            int nextContextIndex = contextIndex + 1;
            var logicalStreams = new List<PdfStream> { context.Stream };
            if (context.PageSequenceId is int logicalPageSequenceId) {
                while (nextContextIndex < contentStreams.Count &&
                       contentStreams[nextContextIndex].PageSequenceId == logicalPageSequenceId) {
                    logicalStreams.Add(contentStreams[nextContextIndex].Stream);
                    nextContextIndex++;
                }
            }
            var colorState = new ContentColorState(context.InitialColorState);
            if (!PdfContentStreamSequenceDecoder.TryDecode(
                    logicalStreams,
                    objects,
                    document.ReadOptions.Limits,
                    enforcePageContentLimit: context.PageSequenceId != null,
                    out string decodedContent)) {
                uninspectable++;
                contextIndex = nextContextIndex - 1;
                continue;
            }

            try {
                bool contextWasUninspectable = colorState.IsIncomplete;
                PdfContentStreamInterpreter.Interpret(
                    decodedContent,
                    document.ReadOptions.Limits.MaxContentOperations,
                    operation => {
                        if (operation.HasInvalidOperands) {
                            contextWasUninspectable = true;
                            return;
                        }

                        switch (operation.Name) {
                            case "q":
                                if (operation.Operands.Count != 0) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                colorState.Stack.Push(new ContentColorStateSnapshot(
                                    colorState.FillUsesDeviceRgb,
                                    colorState.StrokeUsesDeviceRgb,
                                colorState.FillUsesDeviceCmyk,
                                colorState.StrokeUsesDeviceCmyk,
                                colorState.FillUsesDeviceIndependentColor,
                                colorState.StrokeUsesDeviceIndependentColor,
                                colorState.FillComponentCount,
                                colorState.StrokeComponentCount,
                                colorState.FillUsesPattern,
                                colorState.StrokeUsesPattern));
                                break;
                            case "Q":
                                if (operation.Operands.Count != 0) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                if (colorState.Stack.Count > 0) {
                                    ContentColorStateSnapshot restored = colorState.Stack.Pop();
                                    colorState.FillUsesDeviceRgb = restored.FillUsesDeviceRgb;
                                    colorState.StrokeUsesDeviceRgb = restored.StrokeUsesDeviceRgb;
                                    colorState.FillUsesDeviceCmyk = restored.FillUsesDeviceCmyk;
                                    colorState.StrokeUsesDeviceCmyk = restored.StrokeUsesDeviceCmyk;
                                    colorState.FillUsesDeviceIndependentColor = restored.FillUsesDeviceIndependentColor;
                                    colorState.StrokeUsesDeviceIndependentColor = restored.StrokeUsesDeviceIndependentColor;
                                    colorState.FillComponentCount = restored.FillComponentCount;
                                    colorState.StrokeComponentCount = restored.StrokeComponentCount;
                                    colorState.FillUsesPattern = restored.FillUsesPattern;
                                    colorState.StrokeUsesPattern = restored.StrokeUsesPattern;
                                } else {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "rg":
                                if (!HasNumericColorOperands(operation, 3)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                ApplyColorSpaceUsage(
                                    aliases.DefaultRgb ?? ColorSpaceUsage.DeviceRgb,
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref colorState.FillComponentCount,
                                    ref colorState.FillUsesPattern,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "RG":
                                if (!HasNumericColorOperands(operation, 3)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                ApplyColorSpaceUsage(
                                    aliases.DefaultRgb ?? ColorSpaceUsage.DeviceRgb,
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref colorState.StrokeComponentCount,
                                    ref colorState.StrokeUsesPattern,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "k":
                                if (!HasNumericColorOperands(operation, 4)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                ApplyColorSpaceUsage(
                                    aliases.DefaultCmyk ?? ColorSpaceUsage.DeviceCmyk,
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref colorState.FillComponentCount,
                                    ref colorState.FillUsesPattern,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "K":
                                if (!HasNumericColorOperands(operation, 4)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                ApplyColorSpaceUsage(
                                    aliases.DefaultCmyk ?? ColorSpaceUsage.DeviceCmyk,
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref colorState.StrokeComponentCount,
                                    ref colorState.StrokeUsesPattern,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "g":
                                if (!HasNumericColorOperands(operation, 1)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                ApplyColorSpaceUsage(
                                    aliases.DefaultGray ?? ColorSpaceUsage.DeviceGray,
                                    ref colorState.FillUsesDeviceRgb,
                                    ref colorState.FillUsesDeviceCmyk,
                                    ref colorState.FillUsesDeviceIndependentColor,
                                    ref colorState.FillComponentCount,
                                    ref colorState.FillUsesPattern,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "G":
                                if (!HasNumericColorOperands(operation, 1)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                ApplyColorSpaceUsage(
                                    aliases.DefaultGray ?? ColorSpaceUsage.DeviceGray,
                                    ref colorState.StrokeUsesDeviceRgb,
                                    ref colorState.StrokeUsesDeviceCmyk,
                                    ref colorState.StrokeUsesDeviceIndependentColor,
                                    ref colorState.StrokeComponentCount,
                                    ref colorState.StrokeUsesPattern,
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
                                    ref colorState.FillComponentCount,
                                    ref colorState.FillUsesPattern,
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
                                    ref colorState.StrokeComponentCount,
                                    ref colorState.StrokeUsesPattern,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "sc":
                            case "scn":
                                if (!HasSelectedColorOperands(
                                        operation,
                                        colorState.FillComponentCount,
                                        colorState.FillUsesPattern)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                if (colorState.FillUsesDeviceRgb) rgbOperators++;
                                if (colorState.FillUsesDeviceCmyk) cmykOperators++;
                                if (colorState.FillUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "SC":
                            case "SCN":
                                if (!HasSelectedColorOperands(
                                        operation,
                                        colorState.StrokeComponentCount,
                                        colorState.StrokeUsesPattern)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
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
                                    maximumDecodedStreamBytes,
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
                    inlineImageComponentCount: colorSpaceName => ResolveInlineImageComponentCount(
                        new PdfName(colorSpaceName),
                        objects,
                        maximumObjectDepth,
                        maximumDecodedStreamBytes,
                        aliases),
                    maxNestingDepth: document.ReadOptions.Limits.MaxContentNestingDepth,
                    maxOperands: document.ReadOptions.Limits.MaxContentOperands,
                    dispatchInvalidOperations: true,
                    inlineImageArrayComponentCount: colorSpace => ResolveInlineImageComponentCount(
                        colorSpace,
                        objects,
                        maximumObjectDepth,
                        maximumDecodedStreamBytes,
                        aliases));
                bool resourceInspectionIncomplete = false;
                for (int index = contextIndex; index < nextContextIndex; index++) {
                    resourceInspectionIncomplete |= contentStreams[index].ResourceInspectionIncomplete;
                }
                if (contextWasUninspectable || resourceInspectionIncomplete) {
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
            contextIndex = nextContextIndex - 1;
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
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        PdfDictionary? resources = ResolveResourcesDictionary(
            owner,
            objects,
            inheritResources,
            maximumObjectDepth);
        return resources == null
            ? new ColorSpaceAliases()
            : CreateColorSpaceAliases(resources, objects, maximumObjectDepth, maximumDecodedStreamBytes);
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
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        var aliases = new ColorSpaceAliases();
        CollectResourceColorSpaces(
            resources,
            objects,
            aliases,
            maximumObjectDepth,
            maximumDecodedStreamBytes);
        if (resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) &&
            ResolveObject(objects, colorSpacesObject, 0, maximumObjectDepth) is PdfDictionary colorSpaces) {
            aliases.DefaultRgb = ResolveDefaultColorSpaceUsage(colorSpaces, "DefaultRGB", objects, maximumObjectDepth, maximumDecodedStreamBytes);
            aliases.DefaultCmyk = ResolveDefaultColorSpaceUsage(colorSpaces, "DefaultCMYK", objects, maximumObjectDepth, maximumDecodedStreamBytes);
            aliases.DefaultGray = ResolveDefaultColorSpaceUsage(colorSpaces, "DefaultGray", objects, maximumObjectDepth, maximumDecodedStreamBytes);
        }
        return aliases;
    }

    private static ColorSpaceUsage? ResolveDefaultColorSpaceUsage(
        PdfDictionary colorSpaces,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        if (!colorSpaces.Items.TryGetValue(key, out PdfObject? value)) return null;
        int expectedComponents = string.Equals(key, "DefaultRGB", StringComparison.Ordinal) ? 3 :
            string.Equals(key, "DefaultCMYK", StringComparison.Ordinal) ? 4 : 1;
        ColorSpaceUsage usage = ClassifyColorSpace(value, objects, maximumObjectDepth, maximumDecodedStreamBytes);
        return usage.IsKnown &&
               usage.UsesDeviceIndependent &&
               !usage.UsesPattern &&
               usage.ComponentCount == expectedComponents
            ? usage
            : ColorSpaceUsage.Unknown;
    }

    private static ColorSpaceUsage ClassifyColorSpace(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes,
        ColorSpaceAliases? aliases = null,
        bool normalizeInlineImageAbbreviations = false) {
        ColorSpaceUsage usage = ClassifyColorSpaceCore(
            value,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes,
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
        int maximumDecodedStreamBytes,
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
            if (string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal)) return ColorSpaceUsage.DeviceRgb;
            if (string.Equals(colorSpaceName, "DeviceCMYK", StringComparison.Ordinal)) return ColorSpaceUsage.DeviceCmyk;
            if (string.Equals(colorSpaceName, "DeviceGray", StringComparison.Ordinal)) return ColorSpaceUsage.DeviceGray;
            if (string.Equals(colorSpaceName, "Pattern", StringComparison.Ordinal)) return ColorSpaceUsage.Pattern;
            if (aliases?.Named.TryGetValue(colorSpaceName, out ColorSpaceUsage? namedUsage) == true) return namedUsage;
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
                    return array.Items.Count == 2 &&
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is PdfDictionary calGray &&
                        PdfCalibratedColorSpaceSemantics.IsStructurallyValid(
                            "CalGray", calGray, objects, maximumObjectDepth)
                        ? ColorSpaceUsage.DeviceIndependentWithComponents(1)
                        : ColorSpaceUsage.Unknown;
                case "CalRGB":
                    return array.Items.Count == 2 &&
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is PdfDictionary calRgb &&
                        PdfCalibratedColorSpaceSemantics.IsStructurallyValid(
                            "CalRGB", calRgb, objects, maximumObjectDepth)
                        ? ColorSpaceUsage.DeviceIndependentWithComponents(3)
                        : ColorSpaceUsage.Unknown;
                case "Lab":
                    return array.Items.Count == 2 &&
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is PdfDictionary lab &&
                        PdfCalibratedColorSpaceSemantics.IsStructurallyValid(
                            "Lab", lab, objects, maximumObjectDepth)
                        ? ColorSpaceUsage.DeviceIndependentWithComponents(3)
                        : ColorSpaceUsage.Unknown;
                case "ICCBased":
                    if (array.Items.Count != 2 ||
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is not PdfStream profile ||
                        !TryResolveNumber(profile.Dictionary, "N", objects, maximumObjectDepth, out double components) ||
                        (components != 1D && components != 3D && components != 4D) ||
                        !PdfIccProfileCache.TryRead(
                            profile,
                            objects,
                            maximumDecodedStreamBytes,
                            out OfficeIMO.Drawing.OfficeIccColorProfile? parsedProfile) ||
                        parsedProfile == null ||
                        parsedProfile.ComponentCount != (int)components) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ColorSpaceUsage.DeviceIndependentWithComponents((int)components);
                case "Indexed":
                case "I":
                    if (array.Items.Count != 4 ||
                        !TryResolveBoundedInteger(array.Items[2], objects, maximumObjectDepth, 0, 255) ||
                        ResolveObject(objects, array.Items[3], resolvedDepth + 1, maximumObjectDepth) is not (PdfStringObj or PdfStream)) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ClassifyColorSpaceCore(
                        array.Items[1], objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1).WithComponentCount(1);
                case "Separation":
                    if (array.Items.Count != 4 ||
                        ResolveObject(objects, array.Items[1], resolvedDepth + 1, maximumObjectDepth) is not PdfName ||
                        !IsColorSpaceFunction(array.Items[3], objects, maximumObjectDepth, resolvedDepth + 1)) {
                        return ColorSpaceUsage.Unknown;
                    }
                    return ClassifyColorSpaceCore(
                        array.Items[2], objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1).WithComponentCount(1);
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
                        array.Items[2], objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases, normalizeInlineImageAbbreviations,
                        activeArrays, resolvedDepth + 1).WithComponentCount(colorants.Items.Count);
                case "Pattern":
                    if (array.Items.Count == 1) return ColorSpaceUsage.Pattern;
                    if (array.Items.Count != 2) return ColorSpaceUsage.Unknown;
                    ColorSpaceUsage baseUsage = ClassifyColorSpaceCore(
                        array.Items[1], objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases, normalizeInlineImageAbbreviations,
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
        ref int componentCount,
        ref bool usesPattern,
        ref int rgbOperators,
        ref int cmykOperators,
        ref int deviceIndependentColorUses,
        ref bool contextWasUninspectable) {
        usesRgb = usage.UsesDeviceRgb;
        usesCmyk = usage.UsesDeviceCmyk;
        usesDeviceIndependent = usage.UsesDeviceIndependent;
        componentCount = usage.ComponentCount;
        usesPattern = usage.UsesPattern;
        if (!usage.IsKnown) contextWasUninspectable = true;
        if (usesRgb) rgbOperators++;
        if (usesCmyk) cmykOperators++;
        if (usesDeviceIndependent) deviceIndependentColorUses++;
    }

    private static bool HasNumericColorOperands(PdfContentOperation operation, int expectedCount) {
        if (operation.Operands.Count != expectedCount) return false;
        for (int index = 0; index < operation.Operands.Count; index++) {
            if (operation.Operands[index] is not double) return false;
        }
        return true;
    }

    private static bool HasSelectedColorOperands(
        PdfContentOperation operation,
        int componentCount,
        bool usesPattern) {
        if (componentCount < 0) return false;
        int expectedCount = componentCount + (usesPattern ? 1 : 0);
        if (operation.Operands.Count != expectedCount) return false;
        if (usesPattern &&
            ((!string.Equals(operation.Name, "scn", StringComparison.Ordinal) &&
              !string.Equals(operation.Name, "SCN", StringComparison.Ordinal)) ||
             operation.Operands[expectedCount - 1] is not string)) return false;
        for (int index = 0; index < componentCount; index++) {
            if (operation.Operands[index] is not double) return false;
        }
        return true;
    }

    private static ColorSpaceUsage ClassifySelectedColorSpace(
        PdfContentOperation operation,
        ColorSpaceAliases aliases) {
        if (operation.Operands.Count != 1 || operation.Operands[0] is not string colorSpaceName) {
            return ColorSpaceUsage.Unknown;
        }

        ColorSpaceUsage usage;
        if (string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal)) usage = ColorSpaceUsage.DeviceRgb;
        else if (string.Equals(colorSpaceName, "DeviceCMYK", StringComparison.Ordinal)) usage = ColorSpaceUsage.DeviceCmyk;
        else if (string.Equals(colorSpaceName, "DeviceGray", StringComparison.Ordinal)) usage = ColorSpaceUsage.DeviceGray;
        else if (string.Equals(colorSpaceName, "Pattern", StringComparison.Ordinal)) usage = ColorSpaceUsage.Pattern;
        else if (!aliases.Named.TryGetValue(colorSpaceName, out usage!)) usage = ColorSpaceUsage.Unknown;
        if (usage.UsesDeviceRgb && aliases.DefaultRgb != null) usage = usage.ReplaceDeviceRgb(aliases.DefaultRgb);
        if (usage.UsesDeviceCmyk && aliases.DefaultCmyk != null) usage = usage.ReplaceDeviceCmyk(aliases.DefaultCmyk);
        if (usage.UsesDeviceGray && aliases.DefaultGray != null) usage = usage.ReplaceDeviceGray(aliases.DefaultGray);
        return usage;
    }

    private static void CollectResourceColorSpaces(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        ColorSpaceAliases aliases,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            ResolveObject(objects, colorSpacesObject, 0, maximumObjectDepth) is not PdfDictionary colorSpaces) return;

        bool changed;
        do {
            changed = false;
            foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
                ColorSpaceUsage usage = ClassifyColorSpace(entry.Value, objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases);
                if (!usage.IsKnown) continue;
                changed |= AddAliasUsage(entry.Key, usage, aliases);
            }
        } while (changed);

    }

    private static string NormalizeInlineImageColorSpaceName(string name) => name switch {
        "RGB" => "DeviceRGB",
        "CMYK" => "DeviceCMYK",
        "G" => "DeviceGray",
        _ => name
    };

    private static bool AddAliasUsage(string name, ColorSpaceUsage usage, ColorSpaceAliases aliases) {
        bool changed = false;
        if (!aliases.Named.TryGetValue(name, out ColorSpaceUsage? existing) || existing != usage) {
            aliases.Named[name] = usage;
            changed = true;
        }
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
        int maximumDecodedStreamBytes,
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
        usage = ClassifyColorSpace(colorSpace, objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases);
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
        PdfObject? resolved = ResolveObject(objects, value, 0, maximumObjectDepth, out int resolvedDepth);
        if (resolved is PdfName name) return TryClassifyBlendModeName(name.Name, out isNonNormal);
        if (resolved is not PdfArray array || array.Items.Count == 0) return false;

        for (int index = 0; index < array.Items.Count; index++) {
            PdfObject? candidate = ResolveObject(
                objects,
                array.Items[index],
                resolvedDepth + 1,
                maximumObjectDepth);
            if (candidate is not PdfName fallback) return false;
            if (TryClassifyBlendModeName(fallback.Name, out isNonNormal)) return true;
        }
        return false;
    }

    private static int ResolveInlineImageComponentCount(
        PdfObject colorSpace,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes,
        ColorSpaceAliases aliases) {
        ColorSpaceUsage usage = ClassifyColorSpace(
            colorSpace,
            objects,
            maximumObjectDepth,
            maximumDecodedStreamBytes,
            aliases,
            normalizeInlineImageAbbreviations: true);
        return usage.IsKnown && !usage.UsesPattern ? usage.ComponentCount : 0;
    }

    internal static int ResolveInlineImageComponentCountForResources(
        PdfObject colorSpace,
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes) {
        ColorSpaceAliases aliases = resources == null
            ? new ColorSpaceAliases()
            : CreateColorSpaceAliases(resources, objects, maximumObjectDepth, maximumDecodedStreamBytes);
        return ResolveInlineImageComponentCount(colorSpace, objects, maximumObjectDepth, maximumDecodedStreamBytes, aliases);
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
        internal Dictionary<string, ColorSpaceUsage> Named { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Rgb { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Cmyk { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Gray { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Pattern { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> DeviceIndependent { get; } = new(StringComparer.Ordinal);
        internal ColorSpaceUsage? DefaultRgb { get; set; }
        internal ColorSpaceUsage? DefaultCmyk { get; set; }
        internal ColorSpaceUsage? DefaultGray { get; set; }

        internal bool SetEquals(ColorSpaceAliases other) =>
            Named.Count == other.Named.Count && Named.All(entry =>
                other.Named.TryGetValue(entry.Key, out ColorSpaceUsage? usage) && usage == entry.Value) &&
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
        bool UsesDeviceIndependent,
        int ComponentCount) {
        internal static ColorSpaceUsage DeviceRgb { get; } = new(true, true, false, false, false, false, 3);
        internal static ColorSpaceUsage DeviceCmyk { get; } = new(true, false, true, false, false, false, 4);
        internal static ColorSpaceUsage DeviceGray { get; } = new(true, false, false, true, false, false, 1);
        internal static ColorSpaceUsage Pattern { get; } = new(true, false, false, false, true, false, 0);
        internal static ColorSpaceUsage Unknown { get; } = new(false, false, false, false, false, false, -1);

        internal static ColorSpaceUsage DeviceIndependentWithComponents(int componentCount) =>
            new(true, false, false, false, false, true, componentCount);

        internal ColorSpaceUsage ReplaceDeviceRgb(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                replacement.UsesDeviceRgb,
                UsesDeviceCmyk || replacement.UsesDeviceCmyk,
                UsesDeviceGray || replacement.UsesDeviceGray,
                UsesPattern || replacement.UsesPattern,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent,
                ComponentCount);

        internal ColorSpaceUsage ReplaceDeviceCmyk(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                UsesDeviceRgb || replacement.UsesDeviceRgb,
                replacement.UsesDeviceCmyk,
                UsesDeviceGray || replacement.UsesDeviceGray,
                UsesPattern || replacement.UsesPattern,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent,
                ComponentCount);

        internal ColorSpaceUsage ReplaceDeviceGray(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                UsesDeviceRgb || replacement.UsesDeviceRgb,
                UsesDeviceCmyk || replacement.UsesDeviceCmyk,
                replacement.UsesDeviceGray,
                UsesPattern || replacement.UsesPattern,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent,
                ComponentCount);

        internal ColorSpaceUsage WithComponentCount(int componentCount) =>
            IsKnown
                ? new(IsKnown, UsesDeviceRgb, UsesDeviceCmyk, UsesDeviceGray, UsesPattern, UsesDeviceIndependent, componentCount)
                : Unknown;

        internal ColorSpaceUsage WithPattern() =>
            new(IsKnown, UsesDeviceRgb, UsesDeviceCmyk, UsesDeviceGray, true, UsesDeviceIndependent, ComponentCount);
    }

    private sealed class ContentStreamContext {
        internal ContentStreamContext(
            PdfStream stream,
            ColorSpaceAliases aliases,
            PdfDictionary? resources,
            PdfObject? inheritedFontObject,
            int? pageSequenceId,
            ContentColorStateSnapshot? initialColorState) {
            Stream = stream;
            Aliases = aliases;
            Resources = resources;
            InheritedFontObject = inheritedFontObject;
            PageSequenceId = pageSequenceId;
            InitialColorState = initialColorState;
        }

        internal PdfStream Stream { get; }
        internal ColorSpaceAliases Aliases { get; }
        internal PdfDictionary? Resources { get; }
        internal PdfObject? InheritedFontObject { get; }
        internal int? PageSequenceId { get; }
        internal ContentColorStateSnapshot? InitialColorState { get; }
        internal bool ResourceInspectionIncomplete { get; set; }
    }

    private sealed class ContentColorState {
        internal ContentColorState(ContentColorStateSnapshot? initial = null) {
            if (initial is ContentColorStateSnapshot snapshot) Restore(snapshot);
        }

        internal bool FillUsesDeviceRgb;
        internal bool StrokeUsesDeviceRgb;
        internal bool FillUsesDeviceCmyk;
        internal bool StrokeUsesDeviceCmyk;
        internal bool FillUsesDeviceIndependentColor;
        internal bool StrokeUsesDeviceIndependentColor;
        internal int FillComponentCount = 1;
        internal int StrokeComponentCount = 1;
        internal bool FillUsesPattern;
        internal bool StrokeUsesPattern;
        internal bool IsIncomplete;
        internal Stack<ContentColorStateSnapshot> Stack { get; } = new();

        internal ContentColorStateSnapshot Capture() => new(
            FillUsesDeviceRgb,
            StrokeUsesDeviceRgb,
            FillUsesDeviceCmyk,
            StrokeUsesDeviceCmyk,
            FillUsesDeviceIndependentColor,
            StrokeUsesDeviceIndependentColor,
            FillComponentCount,
            StrokeComponentCount,
            FillUsesPattern,
            StrokeUsesPattern);

        internal void Restore(ContentColorStateSnapshot snapshot) {
            FillUsesDeviceRgb = snapshot.FillUsesDeviceRgb;
            StrokeUsesDeviceRgb = snapshot.StrokeUsesDeviceRgb;
            FillUsesDeviceCmyk = snapshot.FillUsesDeviceCmyk;
            StrokeUsesDeviceCmyk = snapshot.StrokeUsesDeviceCmyk;
            FillUsesDeviceIndependentColor = snapshot.FillUsesDeviceIndependentColor;
            StrokeUsesDeviceIndependentColor = snapshot.StrokeUsesDeviceIndependentColor;
            FillComponentCount = snapshot.FillComponentCount;
            StrokeComponentCount = snapshot.StrokeComponentCount;
            FillUsesPattern = snapshot.FillUsesPattern;
            StrokeUsesPattern = snapshot.StrokeUsesPattern;
        }
    }

    private readonly record struct ContentColorStateSnapshot(
        bool FillUsesDeviceRgb,
        bool StrokeUsesDeviceRgb,
        bool FillUsesDeviceCmyk,
        bool StrokeUsesDeviceCmyk,
        bool FillUsesDeviceIndependentColor,
        bool StrokeUsesDeviceIndependentColor,
        int FillComponentCount,
        int StrokeComponentCount,
        bool FillUsesPattern,
        bool StrokeUsesPattern);

    private sealed record ReachableResourceCollection(int TransparencyGroupCount);

    private sealed record ImageContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);

    private sealed record ShadingContext(PdfDictionary Dictionary, PdfStream? Stream, ColorSpaceAliases Aliases);
}
