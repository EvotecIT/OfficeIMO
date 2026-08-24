using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static class PdfPrintProductionColorInspector {
    internal static PdfPrintProductionColorEvidence Inspect(PdfReadDocument document) {
        Guard.NotNull(document, nameof(document));
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        var contentStreams = new List<ContentStreamContext>();
        var graphicsStateDictionaries = new HashSet<PdfDictionary>();
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
            PdfDictionary? dictionary = indirect.Value switch {
                PdfDictionary value => value,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (dictionary == null) continue;

            string? type = ResolveName(dictionary.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null, objects);
            string? subtype = ResolveName(dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null, objects);
            if (indirect.Value is PdfStream candidate &&
                (string.Equals(subtype, "Form", StringComparison.Ordinal) ||
                 dictionary.Items.ContainsKey("PatternType"))) {
                AddContentStream(
                    candidate,
                    ResolveColorSpaceAliases(dictionary, objects, inheritResources: false),
                    contentStreams);
            }

            if (string.Equals(type, "Page", StringComparison.Ordinal) &&
                dictionary.Items.TryGetValue("Contents", out PdfObject? contents)) {
                CollectStreams(
                    contents,
                    objects,
                    ResolveColorSpaceAliases(dictionary, objects, inheritResources: true),
                    contentStreams);
            }

            CollectNestedDictionaryEvidence(
                dictionary,
                objects,
                graphicsStateDictionaries,
                inspectedDictionaries,
                ref transparencyGroups);

            if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                PdfObject? colorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                    ? colorSpaceObject
                    : null;
                if (ContainsColorSpace(colorSpace, "DeviceRGB", objects)) rgbImages++;
                if (ContainsColorSpace(colorSpace, "DeviceCMYK", objects)) cmykImages++;
                if (ContainsDeviceIndependentColorSpace(colorSpace, objects)) deviceIndependentColorUses++;
                if (dictionary.Items.ContainsKey("SMask")) transparentImages++;
            }

            if (dictionary.Items.ContainsKey("ShadingType")) {
                PdfObject? colorSpace = dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)
                    ? colorSpaceObject
                    : null;
                if (ContainsColorSpace(colorSpace, "DeviceRGB", objects)) rgbShadings++;
                if (ContainsColorSpace(colorSpace, "DeviceCMYK", objects)) cmykShadings++;
                if (ContainsDeviceIndependentColorSpace(colorSpace, objects)) deviceIndependentColorUses++;
            }

        }

        foreach (PdfDictionary graphicsState in graphicsStateDictionaries) {
            if (IsNonOpaqueGraphicsState(graphicsState, objects)) nonOpaqueStates++;
        }

        int rgbOperators = 0;
        int cmykOperators = 0;
        int uninspectable = 0;
        foreach (ContentStreamContext context in contentStreams) {
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
                            if (ContainsColorSpace(inlineColorSpace, "DeviceRGB", objects, aliases.Rgb)) rgbImages++;
                            if (ContainsColorSpace(inlineColorSpace, "DeviceCMYK", objects)) cmykImages++;
                            if (ContainsDeviceIndependentColorSpace(inlineColorSpace, objects, aliases.DeviceIndependent)) {
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
        List<ContentStreamContext> streams) {
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, value);
        if (resolved is PdfStream stream) {
            AddContentStream(stream, aliases, streams);
        } else if (resolved is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                CollectStreams(array.Items[index], objects, aliases, streams);
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
        bool inheritResources) {
        var aliases = new ColorSpaceAliases();
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? current = owner;
        while (current != null && visited.Add(current)) {
            if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                PdfObjectLookup.ResolveChain(objects, resourcesObject) is PdfDictionary resources) {
                CollectResourceColorSpaces(resources, objects, aliases.Rgb, aliases.DeviceIndependent);
                break;
            }

            if (!inheritResources ||
                !current.Items.TryGetValue("Parent", out PdfObject? parentObject) ||
                PdfObjectLookup.ResolveChain(objects, parentObject) is not PdfDictionary parent) {
                break;
            }

            current = parent;
        }

        return aliases;
    }

    private static void CollectResourceColorSpaces(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<string> rgbAliases,
        HashSet<string> deviceIndependentAliases) {
        if (!dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            PdfObjectLookup.ResolveChain(objects, colorSpacesObject) is not PdfDictionary colorSpaces) return;

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (ContainsColorSpace(entry.Value, "DeviceRGB", objects)) rgbAliases.Add(entry.Key);
            if (ContainsDeviceIndependentColorSpace(entry.Value, objects)) deviceIndependentAliases.Add(entry.Key);
        }
    }

    private static void CollectNestedDictionaryEvidence(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> graphicsStates,
        HashSet<PdfDictionary> inspected,
        ref int transparencyGroups) {
        if (!inspected.Add(dictionary)) return;

        CollectGraphicsStates(dictionary, objects, graphicsStates);
        if (string.Equals(
                ResolveName(dictionary.Items.TryGetValue("Type", out PdfObject? type) ? type : null, objects),
                "ExtGState",
                StringComparison.Ordinal)) {
            graphicsStates.Add(dictionary);
        }
        if (string.Equals(
                ResolveName(dictionary.Items.TryGetValue("S", out PdfObject? subtype) ? subtype : null, objects),
                "Transparency",
                StringComparison.Ordinal)) {
            transparencyGroups++;
        }

        foreach (PdfObject value in dictionary.Items.Values) {
            CollectNestedDictionaryValue(
                value,
                objects,
                graphicsStates,
                inspected,
                ref transparencyGroups);
        }
    }

    private static void CollectNestedDictionaryValue(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> graphicsStates,
        HashSet<PdfDictionary> inspected,
        ref int transparencyGroups) {
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, value);
        if (resolved is PdfDictionary dictionary) {
            CollectNestedDictionaryEvidence(
                dictionary,
                objects,
                graphicsStates,
                inspected,
                ref transparencyGroups);
        } else if (resolved is PdfStream stream) {
            CollectNestedDictionaryEvidence(
                stream.Dictionary,
                objects,
                graphicsStates,
                inspected,
                ref transparencyGroups);
        } else if (resolved is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                CollectNestedDictionaryValue(
                    array.Items[index],
                    objects,
                    graphicsStates,
                    inspected,
                    ref transparencyGroups);
            }
        }
    }

    private static void CollectGraphicsStates(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> graphicsStates) {
        if (!dictionary.Items.TryGetValue("ExtGState", out PdfObject? graphicsStatesObject) ||
            PdfObjectLookup.ResolveChain(objects, graphicsStatesObject) is not PdfDictionary resources) return;

        foreach (PdfObject value in resources.Items.Values) {
            if (PdfObjectLookup.ResolveChain(objects, value) is PdfDictionary graphicsState) {
                graphicsStates.Add(graphicsState);
            }
        }
    }

    private static bool ContainsColorSpace(
        PdfObject? value,
        string expected,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<string>? aliases = null) {
        PdfObject? resolved = value == null ? null : PdfObjectLookup.ResolveChain(objects, value);
        if (resolved is PdfName name) {
            return string.Equals(name.Name, expected, StringComparison.Ordinal) ||
                expected == "DeviceRGB" && aliases?.Contains(name.Name) == true;
        }

        if (resolved is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                if (ContainsColorSpace(array.Items[index], expected, objects, aliases)) return true;
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
        HashSet<string>? aliases = null) {
        PdfObject? resolved = value == null ? null : PdfObjectLookup.ResolveChain(objects, value);
        if (resolved is PdfName name) {
            return IsDeviceIndependentColorSpaceName(name.Name) || aliases?.Contains(name.Name) == true;
        }
        if (resolved is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                if (ContainsDeviceIndependentColorSpace(array.Items[index], objects, aliases)) return true;
            }
        }
        return false;
    }

    private static bool IsDeviceIndependentColorSpaceName(string name) =>
        string.Equals(name, "CalGray", StringComparison.Ordinal) ||
        string.Equals(name, "CalRGB", StringComparison.Ordinal) ||
        string.Equals(name, "Lab", StringComparison.Ordinal) ||
        string.Equals(name, "ICCBased", StringComparison.Ordinal);

    private static string? ResolveName(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) =>
        value == null ? null : (PdfObjectLookup.ResolveChain(objects, value) as PdfName)?.Name;

    private static bool IsNonOpaqueGraphicsState(PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        if (TryResolveNumber(dictionary, "ca", objects, out double fillAlpha) && fillAlpha != 1D) return true;
        if (TryResolveNumber(dictionary, "CA", objects, out double strokeAlpha) && strokeAlpha != 1D) return true;
        if (dictionary.Items.TryGetValue("BM", out PdfObject? blendObject) &&
            HasNonNormalBlendMode(blendObject, objects)) return true;
        if (!dictionary.Items.TryGetValue("SMask", out PdfObject? softMask)) return false;
        return !string.Equals(ResolveName(softMask, objects), "None", StringComparison.Ordinal);
    }

    private static bool HasNonNormalBlendMode(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects) {
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, value);
        if (resolved is PdfName name) {
            return !string.Equals(name.Name, "Normal", StringComparison.Ordinal);
        }

        if (resolved is PdfArray array) {
            if (array.Items.Count == 0) return true;
            for (int index = 0; index < array.Items.Count; index++) {
                if (HasNonNormalBlendMode(array.Items[index], objects)) return true;
            }

            return false;
        }

        return true;
    }

    private static bool TryResolveNumber(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        out double value) {
        value = 0D;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
            PdfObjectLookup.ResolveChain(objects, candidate) is not PdfNumber number) {
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
