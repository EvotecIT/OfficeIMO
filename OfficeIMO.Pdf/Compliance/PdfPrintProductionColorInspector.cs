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
                    maximumObjectDepth)) uninspectable++;
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

        foreach (ImageContext context in imageContexts) imageDictionaries.Add(context.Dictionary);
        foreach (ShadingContext context in shadingContexts) shadingDictionaries.Add(context.Dictionary);
        foreach (PdfDictionary image in imageDictionaries) {
            if (image.Items.TryGetValue("SMask", out PdfObject? softMask) &&
                !string.Equals(ResolveName(softMask, objects, maximumObjectDepth), "None", StringComparison.Ordinal)) {
                transparentImages++;
            }
        }

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
        foreach (ContentStreamContext context in contentStreams) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfStream stream = context.Stream;
            ColorSpaceAliases aliases = context.Aliases;
            if (!StreamDecoder.TryDecode(
                    stream.Dictionary,
                    stream.Data,
                    document.ReadOptions.Limits.MaxDecodedStreamBytes,
                    out byte[] decoded,
                    objects)) {
                uninspectable++;
                continue;
            }

            try {
                bool fillUsesDeviceRgb = false;
                bool strokeUsesDeviceRgb = false;
                bool fillUsesDeviceCmyk = false;
                bool strokeUsesDeviceCmyk = false;
                bool fillUsesDeviceIndependentColor = false;
                bool strokeUsesDeviceIndependentColor = false;
                var colorSpaceStack = new Stack<(
                    bool FillUsesDeviceRgb,
                    bool StrokeUsesDeviceRgb,
                    bool FillUsesDeviceCmyk,
                    bool StrokeUsesDeviceCmyk,
                    bool FillUsesDeviceIndependentColor,
                    bool StrokeUsesDeviceIndependentColor)>();
                bool contextWasUninspectable = false;
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
                                colorSpaceStack.Push((
                                    fillUsesDeviceRgb,
                                    strokeUsesDeviceRgb,
                                    fillUsesDeviceCmyk,
                                    strokeUsesDeviceCmyk,
                                    fillUsesDeviceIndependentColor,
                                    strokeUsesDeviceIndependentColor));
                                break;
                            case "Q":
                                if (colorSpaceStack.Count > 0) {
                                    (
                                        fillUsesDeviceRgb,
                                        strokeUsesDeviceRgb,
                                        fillUsesDeviceCmyk,
                                        strokeUsesDeviceCmyk,
                                        fillUsesDeviceIndependentColor,
                                        strokeUsesDeviceIndependentColor) = colorSpaceStack.Pop();
                                }
                                break;
                            case "rg":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultRgb ?? ColorSpaceUsage.DeviceRgb,
                                    ref fillUsesDeviceRgb,
                                    ref fillUsesDeviceCmyk,
                                    ref fillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "RG":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultRgb ?? ColorSpaceUsage.DeviceRgb,
                                    ref strokeUsesDeviceRgb,
                                    ref strokeUsesDeviceCmyk,
                                    ref strokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "k":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultCmyk ?? ColorSpaceUsage.DeviceCmyk,
                                    ref fillUsesDeviceRgb,
                                    ref fillUsesDeviceCmyk,
                                    ref fillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "K":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultCmyk ?? ColorSpaceUsage.DeviceCmyk,
                                    ref strokeUsesDeviceRgb,
                                    ref strokeUsesDeviceCmyk,
                                    ref strokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "g":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultGray ?? ColorSpaceUsage.DeviceGray,
                                    ref fillUsesDeviceRgb,
                                    ref fillUsesDeviceCmyk,
                                    ref fillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "G":
                                ApplyColorSpaceUsage(
                                    aliases.DefaultGray ?? ColorSpaceUsage.DeviceGray,
                                    ref strokeUsesDeviceRgb,
                                    ref strokeUsesDeviceCmyk,
                                    ref strokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "cs":
                                ApplyColorSpaceUsage(
                                    ClassifySelectedColorSpace(operation, aliases),
                                    ref fillUsesDeviceRgb,
                                    ref fillUsesDeviceCmyk,
                                    ref fillUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "CS":
                                ApplyColorSpaceUsage(
                                    ClassifySelectedColorSpace(operation, aliases),
                                    ref strokeUsesDeviceRgb,
                                    ref strokeUsesDeviceCmyk,
                                    ref strokeUsesDeviceIndependentColor,
                                    ref rgbOperators,
                                    ref cmykOperators,
                                    ref deviceIndependentColorUses,
                                    ref contextWasUninspectable);
                                break;
                            case "sc":
                            case "scn":
                                if (fillUsesDeviceRgb) rgbOperators++;
                                if (fillUsesDeviceCmyk) cmykOperators++;
                                if (fillUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "SC":
                            case "SCN":
                                if (strokeUsesDeviceRgb) rgbOperators++;
                                if (strokeUsesDeviceCmyk) cmykOperators++;
                                if (strokeUsesDeviceIndependentColor) deviceIndependentColorUses++;
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
                if (contextWasUninspectable || context.ResourceInspectionIncomplete) uninspectable++;
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is PdfReadLimitException ||
                exception is FormatException) {
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
        int maximumObjectDepth) {
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
                AddContentStream(stream, aliases, resources, streams);
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
        int maximumObjectDepth) =>
        colorSpaces.Items.TryGetValue(key, out PdfObject? value)
            ? ClassifyColorSpace(value, objects, maximumObjectDepth)
            : null;

    private static ColorSpaceUsage ClassifyColorSpace(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        ColorSpaceAliases? aliases = null,
        bool normalizeInlineImageAbbreviations = false) {
        bool usesRgb = ContainsColorSpace(
            value,
            "DeviceRGB",
            objects,
            maximumObjectDepth,
            aliases?.Rgb,
            normalizeInlineImageAbbreviations);
        bool usesCmyk = ContainsColorSpace(
            value,
            "DeviceCMYK",
            objects,
            maximumObjectDepth,
            aliases?.Cmyk,
            normalizeInlineImageAbbreviations);
        bool usesGray = ContainsColorSpace(
            value,
            "DeviceGray",
            objects,
            maximumObjectDepth,
            aliases?.Gray,
            normalizeInlineImageAbbreviations);
        bool usesPattern = ContainsColorSpace(
            value,
            "Pattern",
            objects,
            maximumObjectDepth,
            aliases?.Pattern,
            normalizeInlineImageAbbreviations);
        bool usesDeviceIndependent = ContainsDeviceIndependentColorSpace(
            value,
            objects,
            maximumObjectDepth,
            aliases?.DeviceIndependent);

        ColorSpaceUsage usage = new(
            usesRgb || usesCmyk || usesGray || usesPattern || usesDeviceIndependent,
            usesRgb,
            usesCmyk,
            usesDeviceIndependent);
        if (aliases == null) return usage;

        if (usesRgb && aliases.DefaultRgb != null) usage = usage.ReplaceDeviceRgb(aliases.DefaultRgb);
        if (usesCmyk && aliases.DefaultCmyk != null) usage = usage.ReplaceDeviceCmyk(aliases.DefaultCmyk);
        if (usesGray && aliases.DefaultGray != null) usage = usage.Combine(aliases.DefaultGray);
        return usage;
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
            return new ColorSpaceUsage(false, false, false, false);
        }

        bool usesRgb = string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal) ||
            aliases.Rgb.Contains(colorSpaceName);
        bool usesCmyk = string.Equals(colorSpaceName, "DeviceCMYK", StringComparison.Ordinal) ||
            aliases.Cmyk.Contains(colorSpaceName);
        bool usesGray = string.Equals(colorSpaceName, "DeviceGray", StringComparison.Ordinal) ||
            aliases.Gray.Contains(colorSpaceName);
        bool usesPattern = string.Equals(colorSpaceName, "Pattern", StringComparison.Ordinal) ||
            aliases.Pattern.Contains(colorSpaceName);
        bool usesDeviceIndependent = IsDeviceIndependentColorSpaceName(colorSpaceName) ||
            aliases.DeviceIndependent.Contains(colorSpaceName);
        ColorSpaceUsage usage = new(
            usesRgb || usesCmyk || usesGray || usesPattern || usesDeviceIndependent,
            usesRgb,
            usesCmyk,
            usesDeviceIndependent);
        if (usesRgb && aliases.DefaultRgb != null) usage = usage.ReplaceDeviceRgb(aliases.DefaultRgb);
        if (usesCmyk && aliases.DefaultCmyk != null) usage = usage.ReplaceDeviceCmyk(aliases.DefaultCmyk);
        if (usesGray && aliases.DefaultGray != null) usage = usage.Combine(aliases.DefaultGray);
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

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (ContainsColorSpace(entry.Value, "DeviceRGB", objects, maximumObjectDepth)) rgbAliases.Add(entry.Key);
            if (ContainsColorSpace(entry.Value, "DeviceCMYK", objects, maximumObjectDepth)) cmykAliases.Add(entry.Key);
            if (ContainsColorSpace(entry.Value, "DeviceGray", objects, maximumObjectDepth)) grayAliases.Add(entry.Key);
            if (ContainsColorSpace(entry.Value, "Pattern", objects, maximumObjectDepth)) patternAliases.Add(entry.Key);
            if (ContainsDeviceIndependentColorSpace(entry.Value, objects, maximumObjectDepth)) deviceIndependentAliases.Add(entry.Key);
        }
    }

    private static bool ContainsColorSpace(
        PdfObject? value,
        string expected,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        HashSet<string>? aliases = null,
        bool normalizeInlineImageAbbreviations = false) {
        if (value == null) return false;
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
                string colorSpaceName = normalizeInlineImageAbbreviations
                    ? NormalizeInlineImageColorSpaceName(name.Name)
                    : name.Name;
                if (string.Equals(colorSpaceName, expected, StringComparison.Ordinal) ||
                    aliases?.Contains(colorSpaceName) == true) return true;
            } else if (resolved is PdfArray array && inspectedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            }
        }

        return false;
    }

    private static string NormalizeInlineImageColorSpaceName(string name) => name switch {
        "RGB" => "DeviceRGB",
        "CMYK" => "DeviceCMYK",
        "G" => "DeviceGray",
        _ => name
    };

    private static bool ContainsDeviceIndependentColorSpace(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        HashSet<string>? aliases = null) {
        if (value == null) return false;
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
                if (IsDeviceIndependentColorSpaceName(name.Name) || aliases?.Contains(name.Name) == true) return true;
            } else if (resolved is PdfArray array && inspectedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            }
        }
        return false;
    }

    private static bool IsDeviceIndependentColorSpaceName(string name) =>
        string.Equals(name, "CalGray", StringComparison.Ordinal) ||
        string.Equals(name, "CalRGB", StringComparison.Ordinal) ||
        string.Equals(name, "Lab", StringComparison.Ordinal) ||
        string.Equals(name, "ICCBased", StringComparison.Ordinal);

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
        if (dictionary.Items.TryGetValue("BM", out PdfObject? blendObject) &&
            HasNonNormalBlendMode(blendObject, objects, maximumObjectDepth)) isNonOpaque = true;
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
        if (!string.Equals(subtypeName, "Transparency", StringComparison.Ordinal)) return true;
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

    private static bool HasNonNormalBlendMode(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
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
                if (!string.Equals(name.Name, "Normal", StringComparison.Ordinal)) return true;
            } else if (resolved is PdfArray array) {
                if (array.Items.Count == 0 || !inspectedArrays.Add(array)) return true;
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            } else {
                return true;
            }
        }
        return false;
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
        bool UsesDeviceIndependent) {
        internal static ColorSpaceUsage DeviceRgb { get; } = new(true, true, false, false);
        internal static ColorSpaceUsage DeviceCmyk { get; } = new(true, false, true, false);
        internal static ColorSpaceUsage DeviceGray { get; } = new(true, false, false, false);

        internal ColorSpaceUsage ReplaceDeviceRgb(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                replacement.UsesDeviceRgb,
                UsesDeviceCmyk || replacement.UsesDeviceCmyk,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent);

        internal ColorSpaceUsage ReplaceDeviceCmyk(ColorSpaceUsage replacement) =>
            new(
                IsKnown && replacement.IsKnown,
                UsesDeviceRgb || replacement.UsesDeviceRgb,
                replacement.UsesDeviceCmyk,
                UsesDeviceIndependent || replacement.UsesDeviceIndependent);

        internal ColorSpaceUsage Combine(ColorSpaceUsage other) =>
            new(
                IsKnown && other.IsKnown,
                UsesDeviceRgb || other.UsesDeviceRgb,
                UsesDeviceCmyk || other.UsesDeviceCmyk,
                UsesDeviceIndependent || other.UsesDeviceIndependent);
    }

    private sealed class ContentStreamContext {
        internal ContentStreamContext(
            PdfStream stream,
            ColorSpaceAliases aliases,
            PdfDictionary? resources,
            PdfObject? inheritedFontObject) {
            Stream = stream;
            Aliases = aliases;
            Resources = resources;
            InheritedFontObject = inheritedFontObject;
        }

        internal PdfStream Stream { get; }
        internal ColorSpaceAliases Aliases { get; }
        internal PdfDictionary? Resources { get; }
        internal PdfObject? InheritedFontObject { get; }
        internal bool ResourceInspectionIncomplete { get; set; }
    }

    private sealed record ReachableResourceCollection(int TransparencyGroupCount);

    private sealed record ImageContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);

    private sealed record ShadingContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);
}
