using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static class PdfPrintProductionColorInspector {
    internal static PdfPrintProductionColorEvidence Inspect(
        PdfReadDocument document,
        System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        int maximumObjectDepth = document.ReadOptions.Limits.MaxObjectNestingDepth;
        var contentStreams = new List<ContentStreamContext>();
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
                CollectStreams(
                    contents,
                    objects,
                    ResolveColorSpaceAliases(dictionary, objects, inheritResources: true, maximumObjectDepth),
                    contentStreams,
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
                PdfObject? colorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                    ? colorSpaceObject
                    : null;
                if (ContainsColorSpace(colorSpace, "DeviceRGB", objects, maximumObjectDepth)) rgbImages++;
                if (ContainsColorSpace(colorSpace, "DeviceCMYK", objects, maximumObjectDepth)) cmykImages++;
                if (ContainsDeviceIndependentColorSpace(colorSpace, objects, maximumObjectDepth)) deviceIndependentColorUses++;
                if (dictionary.Items.ContainsKey("SMask")) transparentImages++;
            }

        }

        foreach (PdfDictionary shading in shadingDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? colorSpace = shading.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                ? colorSpaceObject
                : null;
            if (ContainsColorSpace(colorSpace, "DeviceRGB", objects, maximumObjectDepth)) rgbShadings++;
            if (ContainsColorSpace(colorSpace, "DeviceCMYK", objects, maximumObjectDepth)) cmykShadings++;
            if (ContainsDeviceIndependentColorSpace(colorSpace, objects, maximumObjectDepth)) deviceIndependentColorUses++;
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
                bool fillUsesDeviceIndependentColor = false;
                bool strokeUsesDeviceIndependentColor = false;
                var colorSpaceStack = new Stack<(
                    bool FillUsesDeviceRgb,
                    bool StrokeUsesDeviceRgb,
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
                                    fillUsesDeviceIndependentColor,
                                    strokeUsesDeviceIndependentColor));
                                break;
                            case "Q":
                                if (colorSpaceStack.Count > 0) {
                                    (
                                        fillUsesDeviceRgb,
                                        strokeUsesDeviceRgb,
                                        fillUsesDeviceIndependentColor,
                                        strokeUsesDeviceIndependentColor) = colorSpaceStack.Pop();
                                }
                                break;
                            case "rg":
                                fillUsesDeviceRgb = true;
                                fillUsesDeviceIndependentColor = false;
                                rgbOperators++;
                                break;
                            case "RG":
                                strokeUsesDeviceRgb = true;
                                strokeUsesDeviceIndependentColor = false;
                                rgbOperators++;
                                break;
                            case "k":
                                fillUsesDeviceRgb = false;
                                fillUsesDeviceIndependentColor = false;
                                cmykOperators++;
                                break;
                            case "K":
                                strokeUsesDeviceRgb = false;
                                strokeUsesDeviceIndependentColor = false;
                                cmykOperators++;
                                break;
                            case "g":
                                fillUsesDeviceRgb = false;
                                fillUsesDeviceIndependentColor = false;
                                break;
                            case "G":
                                strokeUsesDeviceRgb = false;
                                strokeUsesDeviceIndependentColor = false;
                                break;
                            case "cs":
                                fillUsesDeviceRgb = UsesDeviceRgbColorSpace(operation, aliases.Rgb);
                                fillUsesDeviceIndependentColor = UsesDeviceIndependentColorSpace(
                                    operation,
                                    aliases.DeviceIndependent);
                                if (fillUsesDeviceRgb) rgbOperators++;
                                if (fillUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "CS":
                                strokeUsesDeviceRgb = UsesDeviceRgbColorSpace(operation, aliases.Rgb);
                                strokeUsesDeviceIndependentColor = UsesDeviceIndependentColorSpace(
                                    operation,
                                    aliases.DeviceIndependent);
                                if (strokeUsesDeviceRgb) rgbOperators++;
                                if (strokeUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "sc":
                            case "scn":
                                if (fillUsesDeviceRgb) rgbOperators++;
                                if (fillUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                            case "SC":
                            case "SCN":
                                if (strokeUsesDeviceRgb) rgbOperators++;
                                if (strokeUsesDeviceIndependentColor) deviceIndependentColorUses++;
                                break;
                        }

                        if (operation.InlineImage != null &&
                            operation.InlineImage.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? inlineColorSpace)) {
                            if (ContainsColorSpace(inlineColorSpace, "DeviceRGB", objects, maximumObjectDepth, aliases.Rgb)) rgbImages++;
                            if (ContainsColorSpace(inlineColorSpace, "DeviceCMYK", objects, maximumObjectDepth)) cmykImages++;
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

    private static void AddContentStream(
        PdfStream stream,
        ColorSpaceAliases aliases,
        List<ContentStreamContext> streams) {
        for (int index = 0; index < streams.Count; index++) {
            ContentStreamContext existing = streams[index];
            if (ReferenceEquals(existing.Stream, stream) && existing.Aliases.SetEquals(aliases)) return;
        }

        streams.Add(new ContentStreamContext(stream, aliases));
    }

    private static ColorSpaceAliases ResolveColorSpaceAliases(
        PdfDictionary owner,
        Dictionary<int, PdfIndirectObject> objects,
        bool inheritResources,
        int maximumObjectDepth) {
        var aliases = new ColorSpaceAliases();
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
                CollectResourceColorSpaces(
                    resources,
                    objects,
                    aliases.Rgb,
                    aliases.DeviceIndependent,
                    maximumObjectDepth);
                break;
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

        return aliases;
    }

    private static void CollectResourceColorSpaces(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<string> rgbAliases,
        HashSet<string> deviceIndependentAliases,
        int maximumObjectDepth) {
        if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            ResolveObject(objects, colorSpacesObject, 0, maximumObjectDepth) is not PdfDictionary colorSpaces) return;

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (ContainsColorSpace(entry.Value, "DeviceRGB", objects, maximumObjectDepth)) rgbAliases.Add(entry.Key);
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
                if (string.Equals(name.Name, expected, StringComparison.Ordinal) ||
                    expected == "DeviceRGB" && aliases?.Contains(name.Name) == true) return true;
            } else if (resolved is PdfArray array && inspectedArrays.Add(array)) {
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    pending.Push((array.Items[index], resolvedDepth + 1));
                }
            }
        }

        return false;
    }

    private static bool UsesDeviceRgbColorSpace(
        PdfContentOperation operation,
        HashSet<string> rgbAliases) =>
        operation.Operands.Count > 0 &&
        operation.Operands[operation.Operands.Count - 1] is string colorSpaceName &&
        (string.Equals(colorSpaceName, "DeviceRGB", StringComparison.Ordinal) || rgbAliases.Contains(colorSpaceName));

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
        internal HashSet<string> DeviceIndependent { get; } = new(StringComparer.Ordinal);

        internal bool SetEquals(ColorSpaceAliases other) =>
            Rgb.SetEquals(other.Rgb) && DeviceIndependent.SetEquals(other.DeviceIndependent);
    }

    private sealed record ContentStreamContext(PdfStream Stream, ColorSpaceAliases Aliases);
}
