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
        var inspectedDictionaries = new HashSet<PdfDictionary>();
        int rgbImages = 0;
        int cmykImages = 0;
        int rgbShadings = 0;
        int cmykShadings = 0;
        int deviceIndependentColorUses = 0;
        int transparentImages = 0;
        int nonOpaqueStates = 0;
        int transparencyGroups = 0;

        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary? dictionary = indirect.Value switch {
                PdfDictionary value => value,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (dictionary == null) continue;

            string? type = ResolveName(
                dictionary.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null,
                objects,
                maximumObjectDepth);
            string? subtype = ResolveName(
                dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null,
                objects,
                maximumObjectDepth);
            if (indirect.Value is PdfStream candidate &&
                (string.Equals(subtype, "Form", StringComparison.Ordinal) ||
                 dictionary.Items.ContainsKey("PatternType"))) {
                AddContentStream(
                    candidate,
                    ResolveColorSpaceAliases(dictionary, objects, inheritResources: false, maximumObjectDepth),
                    contentStreams);
            }

            if (string.Equals(subtype, "Type3", StringComparison.Ordinal) &&
                dictionary.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) &&
                ResolveObject(objects, charProcsObject, 0, maximumObjectDepth) is PdfDictionary charProcs) {
                ColorSpaceAliases aliases = ResolveColorSpaceAliases(
                    dictionary,
                    objects,
                    inheritResources: false,
                    maximumObjectDepth);
                foreach (PdfObject charProc in charProcs.Items.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    CollectStreams(charProc, objects, aliases, contentStreams, maximumObjectDepth);
                }
            }

            if (string.Equals(type, "Page", StringComparison.Ordinal) &&
                dictionary.Items.TryGetValue("Contents", out PdfObject? contents)) {
                ColorSpaceAliases pageAliases = ResolveColorSpaceAliases(
                    dictionary,
                    objects,
                    inheritResources: true,
                    maximumObjectDepth);
                CollectStreams(
                    contents,
                    objects,
                    pageAliases,
                    contentStreams,
                    maximumObjectDepth);
                PdfDictionary? pageResources = ResolveResourcesDictionary(
                    dictionary,
                    objects,
                    inheritResources: true,
                    maximumObjectDepth);
                if (pageResources != null) {
                    CollectResourceFormStreams(
                        pageResources,
                        objects,
                        pageAliases,
                        contentStreams,
                        imageContexts,
                        shadingContexts,
                        maximumObjectDepth);
                }
            }

            PdfDictionary? directResources = ResolveResourcesDictionary(
                dictionary,
                objects,
                inheritResources: false,
                maximumObjectDepth);
            if (directResources != null) {
                CollectResourceFormStreams(
                    directResources,
                    objects,
                    CreateColorSpaceAliases(directResources, objects, maximumObjectDepth),
                    contentStreams,
                    imageContexts,
                    shadingContexts,
                    maximumObjectDepth);
            }

            CollectNestedDictionaryEvidence(
                dictionary,
                objects,
                graphicsStateDictionaries,
                shadingDictionaries,
                inspectedDictionaries,
                ref transparencyGroups,
                maximumObjectDepth);

            if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                imageDictionaries.Add(dictionary);
                if (dictionary.Items.TryGetValue("SMask", out PdfObject? softMask) &&
                    !string.Equals(ResolveName(softMask, objects, maximumObjectDepth), "None", StringComparison.Ordinal)) {
                    transparentImages++;
                }
            }

        }

        foreach (PdfDictionary image in imageDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? colorSpace = image.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                ? colorSpaceObject
                : null;
            bool usesRgb = ContainsColorSpace(colorSpace, "DeviceRGB", objects, maximumObjectDepth);
            bool usesCmyk = ContainsColorSpace(colorSpace, "DeviceCMYK", objects, maximumObjectDepth);
            bool usesDeviceIndependentColor = ContainsDeviceIndependentColorSpace(
                colorSpace,
                objects,
                maximumObjectDepth);
            foreach (ImageContext context in imageContexts) {
                if (!ReferenceEquals(context.Dictionary, image)) continue;
                usesRgb |= ContainsColorSpace(
                    colorSpace,
                    "DeviceRGB",
                    objects,
                    maximumObjectDepth,
                    context.Aliases.Rgb);
                usesCmyk |= ContainsColorSpace(
                    colorSpace,
                    "DeviceCMYK",
                    objects,
                    maximumObjectDepth,
                    context.Aliases.Cmyk);
                usesDeviceIndependentColor |= ContainsDeviceIndependentColorSpace(
                    colorSpace,
                    objects,
                    maximumObjectDepth,
                    context.Aliases.DeviceIndependent);
            }

            if (usesRgb) rgbImages++;
            if (usesCmyk) cmykImages++;
            if (usesDeviceIndependentColor) deviceIndependentColorUses++;
        }

        foreach (PdfDictionary shading in shadingDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? colorSpace = shading.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                ? colorSpaceObject
                : null;
            bool usesRgb = ContainsColorSpace(colorSpace, "DeviceRGB", objects, maximumObjectDepth);
            bool usesCmyk = ContainsColorSpace(colorSpace, "DeviceCMYK", objects, maximumObjectDepth);
            bool usesDeviceIndependentColor = ContainsDeviceIndependentColorSpace(
                colorSpace,
                objects,
                maximumObjectDepth);
            foreach (ShadingContext context in shadingContexts) {
                if (!ReferenceEquals(context.Dictionary, shading)) continue;
                usesRgb |= ContainsColorSpace(
                    colorSpace,
                    "DeviceRGB",
                    objects,
                    maximumObjectDepth,
                    context.Aliases.Rgb);
                usesCmyk |= ContainsColorSpace(
                    colorSpace,
                    "DeviceCMYK",
                    objects,
                    maximumObjectDepth,
                    context.Aliases.Cmyk);
                usesDeviceIndependentColor |= ContainsDeviceIndependentColorSpace(
                    colorSpace,
                    objects,
                    maximumObjectDepth,
                    context.Aliases.DeviceIndependent);
            }
            if (usesRgb) rgbShadings++;
            if (usesCmyk) cmykShadings++;
            if (usesDeviceIndependentColor) deviceIndependentColorUses++;
        }

        foreach (PdfDictionary graphicsState in graphicsStateDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsNonOpaqueGraphicsState(graphicsState, objects, maximumObjectDepth)) nonOpaqueStates++;
        }

        int rgbOperators = 0;
        int cmykOperators = 0;
        int uninspectable = 0;
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
                PdfContentStreamInterpreter.Interpret(
                    PdfEncoding.Latin1GetString(decoded),
                    document.ReadOptions.Limits.MaxContentOperations,
                    operation => {
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
                                fillUsesDeviceRgb = true;
                                fillUsesDeviceCmyk = false;
                                fillUsesDeviceIndependentColor = false;
                                rgbOperators++;
                                break;
                            case "RG":
                                strokeUsesDeviceRgb = true;
                                strokeUsesDeviceCmyk = false;
                                strokeUsesDeviceIndependentColor = false;
                                rgbOperators++;
                                break;
                            case "k":
                                fillUsesDeviceRgb = false;
                                fillUsesDeviceCmyk = true;
                                fillUsesDeviceIndependentColor = false;
                                cmykOperators++;
                                break;
                            case "K":
                                strokeUsesDeviceRgb = false;
                                strokeUsesDeviceCmyk = true;
                                strokeUsesDeviceIndependentColor = false;
                                cmykOperators++;
                                break;
                            case "g":
                                fillUsesDeviceRgb = false;
                                fillUsesDeviceCmyk = false;
                                fillUsesDeviceIndependentColor = false;
                                break;
                            case "G":
                                strokeUsesDeviceRgb = false;
                                strokeUsesDeviceCmyk = false;
                                strokeUsesDeviceIndependentColor = false;
                                break;
                            case "cs":
                                fillUsesDeviceRgb = UsesDeviceColorSpace(operation, "DeviceRGB", aliases.Rgb);
                                fillUsesDeviceCmyk = UsesDeviceColorSpace(operation, "DeviceCMYK", aliases.Cmyk);
                                fillUsesDeviceIndependentColor = UsesDeviceIndependentColorSpace(
                                    operation,
                                    aliases.DeviceIndependent);
                                if (fillUsesDeviceRgb) rgbOperators++;
                                if (fillUsesDeviceCmyk) cmykOperators++;
                                if (fillUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "CS":
                                strokeUsesDeviceRgb = UsesDeviceColorSpace(operation, "DeviceRGB", aliases.Rgb);
                                strokeUsesDeviceCmyk = UsesDeviceColorSpace(operation, "DeviceCMYK", aliases.Cmyk);
                                strokeUsesDeviceIndependentColor = UsesDeviceIndependentColorSpace(
                                    operation,
                                    aliases.DeviceIndependent);
                                if (strokeUsesDeviceRgb) rgbOperators++;
                                if (strokeUsesDeviceCmyk) cmykOperators++;
                                if (strokeUsesDeviceIndependentColor) deviceIndependentColorUses++;
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

                        if (operation.InlineImage != null &&
                            operation.InlineImage.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? inlineColorSpace)) {
                            if (ContainsColorSpace(
                                    inlineColorSpace,
                                    "DeviceRGB",
                                    objects,
                                    maximumObjectDepth,
                                    aliases.Rgb,
                                    normalizeInlineImageAbbreviations: true)) rgbImages++;
                            if (ContainsColorSpace(
                                    inlineColorSpace,
                                    "DeviceCMYK",
                                    objects,
                                    maximumObjectDepth,
                                    aliases.Cmyk,
                                    normalizeInlineImageAbbreviations: true)) cmykImages++;
                            if (ContainsDeviceIndependentColorSpace(inlineColorSpace, objects, maximumObjectDepth, aliases.DeviceIndependent)) {
                                deviceIndependentColorUses++;
                            }
                        }
                    },
                    maxNestingDepth: document.ReadOptions.Limits.MaxContentNestingDepth,
                    maxOperands: document.ReadOptions.Limits.MaxContentOperands);
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
            deviceIndependentColorUses,
            transparentImages,
            nonOpaqueStates,
            transparencyGroups,
            uninspectable);
    }

    private static void CollectStreams(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        ColorSpaceAliases aliases,
        List<ContentStreamContext> streams,
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
            if (resolved is PdfStream stream) {
                AddContentStream(stream, aliases, streams);
            } else if (resolved is PdfArray array && inspectedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            }
        }
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
            aliases.DeviceIndependent,
            maximumObjectDepth);
        return aliases;
    }

    private static void CollectResourceColorSpaces(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<string> rgbAliases,
        HashSet<string> cmykAliases,
        HashSet<string> deviceIndependentAliases,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            ResolveObject(objects, colorSpacesObject, 0, maximumObjectDepth) is not PdfDictionary colorSpaces) return;

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (ContainsColorSpace(entry.Value, "DeviceRGB", objects, maximumObjectDepth)) rgbAliases.Add(entry.Key);
            if (ContainsColorSpace(entry.Value, "DeviceCMYK", objects, maximumObjectDepth)) cmykAliases.Add(entry.Key);
            if (ContainsDeviceIndependentColorSpace(entry.Value, objects, maximumObjectDepth)) deviceIndependentAliases.Add(entry.Key);
        }
    }

    private static void CollectNestedDictionaryEvidence(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> graphicsStates,
        HashSet<PdfDictionary> shadings,
        HashSet<PdfDictionary> inspected,
        ref int transparencyGroups,
        int maximumObjectDepth) {
        var pending = new Stack<(PdfObject Value, int Depth)>();
        var inspectedArrays = new HashSet<PdfArray>();
        pending.Push((dictionary, 0));
        while (pending.Count > 0) {
            (PdfObject candidate, int depth) = pending.Pop();
            ThrowIfObjectDepthExceeded(depth, maximumObjectDepth);
            PdfObject? resolved = ResolveObject(
                objects,
                candidate,
                depth,
                maximumObjectDepth,
                out int resolvedDepth);
            PdfDictionary? current = resolved switch {
                PdfDictionary resolvedDictionary => resolvedDictionary,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (current != null) {
                if (!inspected.Add(current)) continue;
                CollectGraphicsStates(current, objects, graphicsStates, maximumObjectDepth);
                if (current.Items.ContainsKey("ShadingType")) shadings.Add(current);
                if (string.Equals(
                        ResolveName(
                            current.Items.TryGetValue("Type", out PdfObject? type) ? type : null,
                            objects,
                            maximumObjectDepth),
                        "ExtGState",
                        StringComparison.Ordinal)) {
                    graphicsStates.Add(current);
                }
                if (string.Equals(
                        ResolveName(
                            current.Items.TryGetValue("S", out PdfObject? subtype) ? subtype : null,
                            objects,
                            maximumObjectDepth),
                        "Transparency",
                        StringComparison.Ordinal)) {
                    transparencyGroups++;
                }
                foreach (PdfObject child in current.Items.Values) {
                    pending.Push((child, resolvedDepth + 1));
                }
            } else if (resolved is PdfArray array && inspectedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            }
        }
    }

    private static void CollectGraphicsStates(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> graphicsStates,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue("ExtGState", out PdfObject? graphicsStatesObject) ||
            ResolveObject(objects, graphicsStatesObject, 0, maximumObjectDepth) is not PdfDictionary resources) return;

        foreach (PdfObject value in resources.Items.Values) {
            if (ResolveObject(objects, value, 0, maximumObjectDepth) is PdfDictionary graphicsState) {
                graphicsStates.Add(graphicsState);
            }
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

    private static bool UsesDeviceColorSpace(
        PdfContentOperation operation,
        string expected,
        HashSet<string> aliases) =>
        operation.Operands.Count > 0 &&
        operation.Operands[operation.Operands.Count - 1] is string colorSpaceName &&
        (string.Equals(colorSpaceName, expected, StringComparison.Ordinal) || aliases.Contains(colorSpaceName));

    private static string NormalizeInlineImageColorSpaceName(string name) => name switch {
        "RGB" => "DeviceRGB",
        "CMYK" => "DeviceCMYK",
        "G" => "DeviceGray",
        _ => name
    };

    private static bool UsesDeviceIndependentColorSpace(
        PdfContentOperation operation,
        HashSet<string> aliases) =>
        operation.Operands.Count > 0 &&
        operation.Operands[operation.Operands.Count - 1] is string colorSpaceName &&
        (IsDeviceIndependentColorSpaceName(colorSpaceName) || aliases.Contains(colorSpaceName));

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

    private static bool IsNonOpaqueGraphicsState(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (TryResolveNumber(dictionary, "ca", objects, maximumObjectDepth, out double fillAlpha) && fillAlpha != 1D) return true;
        if (TryResolveNumber(dictionary, "CA", objects, maximumObjectDepth, out double strokeAlpha) && strokeAlpha != 1D) return true;
        if (dictionary.Items.TryGetValue("BM", out PdfObject? blendObject) &&
            HasNonNormalBlendMode(blendObject, objects, maximumObjectDepth)) return true;
        if (!dictionary.Items.TryGetValue("SMask", out PdfObject? softMask)) return false;
        return !string.Equals(ResolveName(softMask, objects, maximumObjectDepth), "None", StringComparison.Ordinal);
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
        internal HashSet<string> DeviceIndependent { get; } = new(StringComparer.Ordinal);

        internal bool SetEquals(ColorSpaceAliases other) =>
            Rgb.SetEquals(other.Rgb) &&
            Cmyk.SetEquals(other.Cmyk) &&
            DeviceIndependent.SetEquals(other.DeviceIndependent);
    }

    private sealed record ContentStreamContext(PdfStream Stream, ColorSpaceAliases Aliases);

    private sealed record ImageContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);

    private sealed record ShadingContext(PdfDictionary Dictionary, ColorSpaceAliases Aliases);
}
