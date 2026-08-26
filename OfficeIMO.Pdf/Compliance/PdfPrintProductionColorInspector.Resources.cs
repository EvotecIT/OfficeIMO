using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static void AddContentStream(
        PdfStream stream,
        ColorSpaceAliases aliases,
        PdfDictionary? resources,
        List<ContentStreamContext> streams) {
        if (ContainsContentStreamContext(streams, stream, aliases, resources)) return;

        streams.Add(new ContentStreamContext(stream, aliases, resources));
    }

    private static bool ContainsContentStreamContext(
        IReadOnlyList<ContentStreamContext> contexts,
        PdfStream stream,
        ColorSpaceAliases aliases,
        PdfDictionary? resources) {
        for (int index = 0; index < contexts.Count; index++) {
            ContentStreamContext existing = contexts[index];
            if (ReferenceEquals(existing.Stream, stream) &&
                existing.Aliases.SetEquals(aliases) &&
                ReferenceEquals(existing.Resources, resources)) return true;
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
            try {
                PdfContentStreamInterpreter.Interpret(
                    PdfEncoding.Latin1GetString(decoded),
                    limits.MaxContentOperations,
                    operation => {
                        cancellationToken.ThrowIfCancellationRequested();
                        switch (operation.Name) {
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
                                        ref transparencyGroups)) {
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
                                    !TryAddType3CharProcs(
                                        fontName,
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
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is PdfReadLimitException ||
                exception is FormatException) {
                // The normal inspection pass records decoding/interpreter failures once.
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
        ref int transparencyGroups) {
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
        AddContentStream(stream, aliases, resources, streams);
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
            !softMask.Items.TryGetValue("G", out PdfObject? groupObject) ||
            ResolveObject(
                objects,
                groupObject,
                softMaskDepth + 1,
                limits.MaxObjectNestingDepth,
                out int groupDepth) is not PdfStream group) return false;
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

    private static bool TryAddType3CharProcs(
        string fontName,
        ContentStreamContext context,
        int contentDepth,
        List<ContentStreamContext> streams,
        List<int> contentDepths,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        ref int transparencyGroups) {
        if (!TryResolveResource(
                context.Resources,
                "Font",
                fontName,
                objects,
                limits.MaxObjectNestingDepth,
                out PdfObject? fontObject) ||
            ResolveObject(
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
        foreach (PdfObject charProcObject in charProcs.Items.Values) {
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
        return charProcs.Items.Count > 0;
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
