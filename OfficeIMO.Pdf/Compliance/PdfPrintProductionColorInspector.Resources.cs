using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static void AddContentStream(
        PdfStream stream,
        ColorSpaceAliases aliases,
        PdfDictionary? resources,
        List<ContentStreamContext> streams,
        PdfObject? inheritedFontObject = null) {
        if (ContainsContentStreamContext(streams, stream, aliases, resources, inheritedFontObject)) return;

        streams.Add(new ContentStreamContext(stream, aliases, resources, inheritedFontObject));
    }

    private static bool ContainsContentStreamContext(
        IReadOnlyList<ContentStreamContext> contexts,
        PdfStream stream,
        ColorSpaceAliases aliases,
        PdfDictionary? resources,
        PdfObject? inheritedFontObject) {
        for (int index = 0; index < contexts.Count; index++) {
            ContentStreamContext existing = contexts[index];
            if (ReferenceEquals(existing.Stream, stream) &&
                existing.Aliases.SetEquals(aliases) &&
                ReferenceEquals(existing.Resources, resources) &&
                ReferenceEquals(existing.InheritedFontObject, inheritedFontObject)) return true;
        }
        return false;
    }

    private static bool AddResolvedShadingContext(
        PdfObject value,
        int depth,
        ColorSpaceAliases aliases,
        Dictionary<int, PdfIndirectObject> objects,
        List<ShadingContext> shadings,
        int maximumObjectDepth) {
        PdfObject? resolved = ResolveObject(objects, value, depth, maximumObjectDepth);
        PdfDictionary? shading = resolved switch {
            PdfDictionary dictionary => dictionary,
            PdfStream stream => stream.Dictionary,
            _ => null
        };
        if (shading == null) return false;
        for (int index = 0; index < shadings.Count; index++) {
            ShadingContext existing = shadings[index];
            if (ReferenceEquals(existing.Dictionary, shading) && existing.Aliases.SetEquals(aliases)) return true;
        }
        shadings.Add(new ShadingContext(shading, aliases));
        return true;
    }

    private static void AddImageContext(
        PdfDictionary image,
        ColorSpaceAliases aliases,
        List<ImageContext> images) {
        for (int index = 0; index < images.Count; index++) {
            ImageContext existing = images[index];
            if (ReferenceEquals(existing.Dictionary, image) && existing.Aliases.SetEquals(aliases)) return;
        }

        images.Add(new ImageContext(image, aliases));
    }

    private static ReachableResourceCollection CollectReachableResourceContexts(
        List<ContentStreamContext> streams,
        int firstContext,
        Dictionary<int, PdfIndirectObject> objects,
        List<ImageContext> images,
        List<ShadingContext> shadings,
        HashSet<PdfDictionary> graphicsStates,
        PdfReadLimits limits,
        System.Threading.CancellationToken cancellationToken) {
        var contentDepths = new List<int>();
        for (int index = firstContext; index < streams.Count; index++) contentDepths.Add(0);

        int transparencyGroups = 0;
        PdfObject? activePageFontObject = null;
        var pageFontStack = new Stack<PdfObject?>();
        for (int localIndex = 0; localIndex < contentDepths.Count; localIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            ContentStreamContext context = streams[firstContext + localIndex];
            if (!StreamDecoder.TryDecode(
                    context.Stream.Dictionary,
                    context.Stream.Data,
                    limits.MaxDecodedStreamBytes,
                    out byte[] decoded,
                    objects)) continue;

            bool contextWasUninspectable = false;
            bool isPageContent = contentDepths[localIndex] == 0;
            PdfObject? activeFontObject = isPageContent ? activePageFontObject : context.InheritedFontObject;
            Stack<PdfObject?> fontStack = isPageContent ? pageFontStack : new Stack<PdfObject?>();
            try {
                PdfContentStreamInterpreter.Interpret(
                    PdfEncoding.Latin1GetString(decoded),
                    limits.MaxContentOperations,
                    operation => {
                        cancellationToken.ThrowIfCancellationRequested();
                        switch (operation.Name) {
                            case "q":
                                fontStack.Push(activeFontObject);
                                break;
                            case "Q":
                                if (fontStack.Count > 0) activeFontObject = fontStack.Pop();
                                break;
                            case "Do":
                                if (operation.HasInvalidOperands ||
                                    operation.Operands.Count != 1 ||
                                    operation.Operands[0] is not string xObjectName ||
                                    !TryResolveResource(
                                        context.Resources,
                                        "XObject",
                                        xObjectName,
                                        objects,
                                        limits.MaxObjectNestingDepth,
                                        out PdfObject? xObject) ||
                                    ResolveObject(
                                        objects,
                                        xObject,
                                        0,
                                        limits.MaxObjectNestingDepth,
                                        out int xObjectDepth) is not PdfStream xObjectStream) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                string? subtype = ResolveName(
                                    xObjectStream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)
                                        ? subtypeObject
                                        : null,
                                    objects,
                                    limits.MaxObjectNestingDepth);
                                if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                                    AddImageContext(xObjectStream.Dictionary, context.Aliases, images);
                                } else if (!string.Equals(subtype, "Form", StringComparison.Ordinal) ||
                                    !AddNestedStream(
                                        xObjectStream,
                                        xObjectDepth,
                                        context.Resources,
                                        context.Aliases,
                                        contentDepths[localIndex] + 1,
                                        streams,
                                        contentDepths,
                                        objects,
                                        limits,
                                        ref transparencyGroups,
                                        activeFontObject)) {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "sh":
                                if (operation.HasInvalidOperands ||
                                    operation.Operands.Count != 1 ||
                                    operation.Operands[0] is not string shadingName ||
                                    !TryResolveResource(
                                        context.Resources,
                                        "Shading",
                                        shadingName,
                                        objects,
                                        limits.MaxObjectNestingDepth,
                                        out PdfObject? shadingObject)) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                if (!AddResolvedShadingContext(
                                        shadingObject!,
                                        0,
                                        context.Aliases,
                                        objects,
                                        shadings,
                                        limits.MaxObjectNestingDepth)) {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "gs":
                                if (operation.HasInvalidOperands ||
                                    operation.Operands.Count != 1 ||
                                    operation.Operands[0] is not string graphicsStateName ||
                                    !TryResolveResource(
                                        context.Resources,
                                        "ExtGState",
                                        graphicsStateName,
                                        objects,
                                        limits.MaxObjectNestingDepth,
                                        out PdfObject? graphicsStateObject) ||
                                    ResolveObject(
                                        objects,
                                        graphicsStateObject,
                                        0,
                                        limits.MaxObjectNestingDepth,
                                        out int graphicsStateDepth) is not PdfDictionary graphicsState) {
                                    contextWasUninspectable = true;
                                    break;
                                }
                                graphicsStates.Add(graphicsState);
                                if (!TryAddSoftMaskStream(
                                        graphicsState,
                                        graphicsStateDepth,
                                        context,
                                        contentDepths[localIndex] + 1,
                                        streams,
                                        contentDepths,
                                        objects,
                                        limits,
                                        ref transparencyGroups)) {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "scn":
                            case "SCN":
                                if (operation.Operands.Count > 0 &&
                                    operation.Operands[operation.Operands.Count - 1] is string patternName &&
                                    !TryAddPatternContext(
                                        patternName,
                                        context,
                                        contentDepths[localIndex] + 1,
                                        streams,
                                        contentDepths,
                                        objects,
                                        shadings,
                                        limits,
                                        ref transparencyGroups)) {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "Tf":
                                if (operation.HasInvalidOperands ||
                                    operation.Operands.Count != 2 ||
                                    operation.Operands[0] is not string fontName ||
                                    !TryResolveResource(
                                        context.Resources,
                                        "Font",
                                        fontName,
                                        objects,
                                        limits.MaxObjectNestingDepth,
                                        out activeFontObject)) {
                                    contextWasUninspectable = true;
                                }
                                break;
                            case "Tj":
                            case "TJ":
                            case "'":
                            case "\"":
                                if (operation.HasInvalidOperands ||
                                    activeFontObject == null ||
                                    !TryAddShownType3CharProcs(
                                        activeFontObject,
                                        operation,
                                        context,
                                        contentDepths[localIndex] + 1,
                                        streams,
                                        contentDepths,
                                        objects,
                                        limits,
                                        ref transparencyGroups)) {
                                    contextWasUninspectable = true;
                                }
                                break;
                        }
                    },
                    maxNestingDepth: limits.MaxContentNestingDepth,
                    maxOperands: limits.MaxContentOperands,
                    dispatchInvalidOperations: true);
                if (isPageContent) activePageFontObject = activeFontObject;
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is PdfReadLimitException ||
                exception is FormatException) {
                // Resource traversal owns reachability. If it cannot resolve a referenced
                // resource, the later color pass cannot recover that missing evidence.
                context.ResourceInspectionIncomplete = true;
                continue;
            }
            if (contextWasUninspectable) context.ResourceInspectionIncomplete = true;
        }

        return new ReachableResourceCollection(transparencyGroups);
    }

    private static bool AddNestedStream(
        PdfStream stream,
        int objectDepth,
        PdfDictionary? inheritedResources,
        ColorSpaceAliases inheritedAliases,
        int contentDepth,
        List<ContentStreamContext> streams,
        List<int> contentDepths,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        ref int transparencyGroups,
        PdfObject? inheritedFontObject = null) {
        if (contentDepth > limits.MaxContentNestingDepth) return false;

        PdfDictionary? resources = inheritedResources;
        ColorSpaceAliases aliases = inheritedAliases;
        if (stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)) {
            if (ResolveObject(
                    objects,
                    resourcesObject,
                    objectDepth + 1,
                    limits.MaxObjectNestingDepth,
                    out _) is not PdfDictionary directResources) return false;
            resources = directResources;
            aliases = CreateColorSpaceAliases(directResources, objects, limits.MaxObjectNestingDepth);
        }

        int count = streams.Count;
        AddContentStream(stream, aliases, resources, streams, inheritedFontObject);
        if (streams.Count == count) return true;
        contentDepths.Add(contentDepth);
        if (IsTransparencyGroup(stream.Dictionary, objects, limits.MaxObjectNestingDepth)) transparencyGroups++;
        return true;
    }

    private static bool TryAddSoftMaskStream(
        PdfDictionary graphicsState,
        int graphicsStateDepth,
        ContentStreamContext context,
        int contentDepth,
        List<ContentStreamContext> streams,
        List<int> contentDepths,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        ref int transparencyGroups) {
        if (!graphicsState.Items.TryGetValue("SMask", out PdfObject? softMaskObject)) return true;
        if (string.Equals(
                ResolveName(softMaskObject, objects, limits.MaxObjectNestingDepth),
                "None",
                StringComparison.Ordinal)) return true;
        if (ResolveObject(
                objects,
                softMaskObject,
                graphicsStateDepth + 1,
                limits.MaxObjectNestingDepth,
                out int softMaskDepth) is not PdfDictionary softMask ||
            ResolveName(
                softMask.Items.TryGetValue("S", out PdfObject? softMaskSubtypeObject)
                    ? softMaskSubtypeObject
                    : null,
                objects,
                limits.MaxObjectNestingDepth) is not string softMaskSubtype ||
            softMaskSubtype is not ("Alpha" or "Luminosity") ||
            !softMask.Items.TryGetValue("G", out PdfObject? groupObject) ||
            ResolveObject(
                objects,
                groupObject,
                softMaskDepth + 1,
                limits.MaxObjectNestingDepth,
                out int groupDepth) is not PdfStream group) return false;
        if (!string.Equals(
                ResolveName(
                    group.Dictionary.Items.TryGetValue("Subtype", out PdfObject? groupSubtypeObject)
                        ? groupSubtypeObject
                        : null,
                    objects,
                    limits.MaxObjectNestingDepth),
                "Form",
                StringComparison.Ordinal) ||
            !TryClassifyTransparencyGroup(
                group.Dictionary,
                context.Aliases,
                objects,
                limits.MaxObjectNestingDepth,
                out bool isTransparencyGroup,
                out _) ||
            !isTransparencyGroup) return false;
        return AddNestedStream(
            group,
            groupDepth,
            context.Resources,
            context.Aliases,
            contentDepth,
            streams,
            contentDepths,
            objects,
            limits,
            ref transparencyGroups);
    }

    private static bool TryAddPatternContext(
        string patternName,
        ContentStreamContext context,
        int contentDepth,
        List<ContentStreamContext> streams,
        List<int> contentDepths,
        Dictionary<int, PdfIndirectObject> objects,
        List<ShadingContext> shadings,
        PdfReadLimits limits,
        ref int transparencyGroups) {
        if (!TryResolveResource(
                context.Resources,
                "Pattern",
                patternName,
                objects,
                limits.MaxObjectNestingDepth,
                out PdfObject? patternObject)) return false;
        PdfObject? resolved = ResolveObject(
            objects,
            patternObject,
            0,
            limits.MaxObjectNestingDepth,
            out int patternDepth);
        PdfDictionary? pattern = resolved switch {
            PdfDictionary dictionary => dictionary,
            PdfStream stream => stream.Dictionary,
            _ => null
        };
        if (pattern == null ||
            !pattern.Items.TryGetValue("PatternType", out PdfObject? patternTypeObject) ||
            ResolveObject(
                objects,
                patternTypeObject,
                patternDepth + 1,
                limits.MaxObjectNestingDepth) is not PdfNumber patternType) return false;
        if (patternType.Value == 1D && resolved is PdfStream tilingPattern) {
            return AddNestedStream(
                tilingPattern,
                patternDepth,
                context.Resources,
                context.Aliases,
                contentDepth,
                streams,
                contentDepths,
                objects,
                limits,
                ref transparencyGroups);
        }
        if (patternType.Value == 2D && pattern.Items.TryGetValue("Shading", out PdfObject? shadingObject)) {
            return AddResolvedShadingContext(
                shadingObject,
                patternDepth + 1,
                context.Aliases,
                objects,
                shadings,
                limits.MaxObjectNestingDepth);
        }
        return false;
    }

    private static bool TryAddShownType3CharProcs(
        PdfObject fontObject,
        PdfContentOperation operation,
        ContentStreamContext context,
        int contentDepth,
        List<ContentStreamContext> streams,
        List<int> contentDepths,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        ref int transparencyGroups) {
        if (ResolveObject(
                objects,
                fontObject,
                0,
                limits.MaxObjectNestingDepth,
                out int fontDepth) is not PdfDictionary font) return false;
        string? subtype = ResolveName(
            font.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null,
            objects,
            limits.MaxObjectNestingDepth);
        if (!string.Equals(subtype, "Type3", StringComparison.Ordinal)) return true;
        if (!font.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) ||
            ResolveObject(
                objects,
                charProcsObject,
                fontDepth + 1,
                limits.MaxObjectNestingDepth,
                out int charProcsDepth) is not PdfDictionary charProcs) return false;
        if (!TryGetType3GlyphNames(font, objects, limits.MaxObjectNestingDepth, out Dictionary<int, string> glyphNames) ||
            !TryGetShownTextBytes(operation, out List<byte[]> shownText)) return false;

        PdfDictionary? resources = context.Resources;
        ColorSpaceAliases aliases = context.Aliases;
        if (font.Items.TryGetValue("Resources", out PdfObject? resourcesObject)) {
            if (ResolveObject(
                    objects,
                    resourcesObject,
                    fontDepth + 1,
                    limits.MaxObjectNestingDepth,
                    out _) is not PdfDictionary fontResources) return false;
            resources = fontResources;
            aliases = CreateColorSpaceAliases(fontResources, objects, limits.MaxObjectNestingDepth);
        }
        var shownGlyphNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (byte[] text in shownText) {
            for (int index = 0; index < text.Length; index++) {
                if (!glyphNames.TryGetValue(text[index], out string? glyphName)) return false;
                shownGlyphNames.Add(glyphName);
            }
        }
        foreach (string glyphName in shownGlyphNames) {
            if (!charProcs.Items.TryGetValue(glyphName, out PdfObject? charProcObject)) return false;
            if (ResolveObject(
                    objects,
                    charProcObject,
                    charProcsDepth + 1,
                    limits.MaxObjectNestingDepth,
                    out int charProcDepth) is not PdfStream charProc ||
                !AddNestedStream(
                    charProc,
                    charProcDepth,
                    resources,
                    aliases,
                    contentDepth,
                    streams,
                    contentDepths,
                    objects,
                    limits,
                    ref transparencyGroups)) return false;
        }
        return true;
    }

    internal static bool TryGetType3GlyphNames(
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out Dictionary<int, string> glyphNames) {
        glyphNames = null!;
        if (!font.Items.TryGetValue("Encoding", out PdfObject? encodingObject)) return false;
        PdfObject? resolvedEncoding = ResolveObject(objects, encodingObject, 0, maximumObjectDepth);
        PdfDictionary? encoding = resolvedEncoding as PdfDictionary;
        string baseEncoding = "StandardEncoding";
        if (resolvedEncoding is PdfName encodingName) {
            baseEncoding = encodingName.Name;
        } else if (encoding == null) {
            return false;
        }
        if (encoding?.Items.TryGetValue("BaseEncoding", out PdfObject? baseEncodingObject) == true) {
            if (ResolveObject(objects, baseEncodingObject, 0, maximumObjectDepth) is not PdfName baseEncodingName) {
                return false;
            }
            baseEncoding = baseEncodingName.Name;
        }
        if (!PdfType3GlyphEncoding.TryCreate(baseEncoding, out glyphNames)) return false;
        if (encoding?.Items.TryGetValue("Differences", out PdfObject? differencesObject) != true) return true;
        if (ResolveObject(objects, differencesObject, 0, maximumObjectDepth) is not PdfArray differences) return false;

        int code = -1;
        foreach (PdfObject item in differences.Items) {
            PdfObject? resolved = ResolveObject(objects, item, 0, maximumObjectDepth);
            if (resolved is PdfNumber number) {
                if (double.IsNaN(number.Value) || double.IsInfinity(number.Value) ||
                    number.Value != Math.Truncate(number.Value) || number.Value < 0D || number.Value > 255D) return false;
                code = (int)number.Value;
            } else if (resolved is PdfName name && code >= 0 && code <= 255) {
                glyphNames[code++] = name.Name;
            } else {
                return false;
            }
        }
        return true;
    }

    internal static bool TryGetShownTextBytes(PdfContentOperation operation, out List<byte[]> shownText) {
        shownText = new List<byte[]>();
        if (string.Equals(operation.Name, "TJ", StringComparison.Ordinal)) {
            if (operation.Operands.Count != 1 || operation.Operands[0] is not List<object> items) return false;
            foreach (object item in items) {
                if (item is byte[] bytes) {
                    shownText.Add(bytes);
                } else if (item is not double value || double.IsNaN(value) || double.IsInfinity(value)) {
                    return false;
                }
            }
            return true;
        }

        int expectedOperandCount = string.Equals(operation.Name, "\"", StringComparison.Ordinal) ? 3 : 1;
        if (operation.Operands.Count != expectedOperandCount ||
            operation.Operands[operation.Operands.Count - 1] is not byte[] text) return false;
        shownText.Add(text);
        return true;
    }

    private static bool IsTransparencyGroup(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        dictionary.Items.TryGetValue("Group", out PdfObject? groupObject) &&
        ResolveObject(objects, groupObject, 0, maximumObjectDepth) is PdfDictionary group &&
        string.Equals(
            ResolveName(
                group.Items.TryGetValue("S", out PdfObject? subtypeObject) ? subtypeObject : null,
                objects,
                maximumObjectDepth),
            "Transparency",
            StringComparison.Ordinal);

    private static bool TryResolveResource(
        PdfDictionary? resources,
        string category,
        string name,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out PdfObject? resource) {
        resource = null;
        return resources != null &&
            resources.Items.TryGetValue(category, out PdfObject? categoryObject) &&
            ResolveObject(objects, categoryObject, 0, maximumObjectDepth) is PdfDictionary entries &&
            entries.Items.TryGetValue(name, out resource);
    }

}
