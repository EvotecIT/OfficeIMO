namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static void AddContentStream(
        PdfStream stream,
        ColorSpaceAliases aliases,
        List<ContentStreamContext> streams) {
        if (ContainsContentStreamContext(streams, stream, aliases)) return;

        streams.Add(new ContentStreamContext(stream, aliases));
    }

    private static bool ContainsContentStreamContext(
        IReadOnlyList<ContentStreamContext> contexts,
        PdfStream stream,
        ColorSpaceAliases aliases) {
        for (int index = 0; index < contexts.Count; index++) {
            ContentStreamContext existing = contexts[index];
            if (ReferenceEquals(existing.Stream, stream) && existing.Aliases.SetEquals(aliases)) return true;
        }
        return false;
    }

    private static void CollectResourceFormStreams(
        PdfDictionary rootResources,
        Dictionary<int, PdfIndirectObject> objects,
        ColorSpaceAliases inheritedAliases,
        List<ContentStreamContext> streams,
        List<ImageContext> images,
        List<ShadingContext> shadings,
        int maximumObjectDepth) {
        var pending = new Stack<(PdfDictionary Resources, ColorSpaceAliases Aliases, int Depth)>();
        var visitedContexts = new List<ContentStreamContext>();
        pending.Push((rootResources, inheritedAliases, 0));
        while (pending.Count > 0) {
            (PdfDictionary resources, ColorSpaceAliases aliases, int depth) = pending.Pop();
            ThrowIfObjectDepthExceeded(depth, maximumObjectDepth);
            CollectResourceShadingContexts(
                resources,
                aliases,
                objects,
                shadings,
                depth,
                maximumObjectDepth);
            CollectResourceType3Streams(
                resources,
                aliases,
                objects,
                streams,
                pending,
                depth,
                maximumObjectDepth);
            CollectResourcePatternStreams(
                resources,
                aliases,
                objects,
                streams,
                visitedContexts,
                pending,
                depth,
                maximumObjectDepth);
            if (resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) &&
                ResolveObject(
                    objects,
                    xObjectsObject,
                    depth + 1,
                    maximumObjectDepth,
                    out int xObjectsDepth) is PdfDictionary xObjects) {
                foreach (PdfObject xObject in xObjects.Items.Values) {
                    PdfObject? resolved = ResolveObject(
                        objects,
                        xObject,
                        xObjectsDepth + 1,
                        maximumObjectDepth,
                        out int formDepth);
                    if (resolved is PdfStream form) {
                        string? subtype = ResolveName(
                            form.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)
                                ? subtypeObject
                                : null,
                            objects,
                            maximumObjectDepth);
                        if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                            AddImageContext(form.Dictionary, aliases, images);
                        } else {
                            AddInvokedFormContext(
                                form,
                                formDepth,
                                resources,
                                aliases,
                                objects,
                                streams,
                                visitedContexts,
                                pending,
                                maximumObjectDepth);
                        }
                    }
                }
            }

            if (!resources.Items.TryGetValue("ExtGState", out PdfObject? graphicsStatesObject) ||
                ResolveObject(
                    objects,
                    graphicsStatesObject,
                    depth + 1,
                    maximumObjectDepth,
                    out int graphicsStatesDepth) is not PdfDictionary graphicsStates) continue;
            foreach (PdfObject graphicsStateObject in graphicsStates.Items.Values) {
                if (ResolveObject(
                        objects,
                        graphicsStateObject,
                        graphicsStatesDepth + 1,
                        maximumObjectDepth,
                        out int graphicsStateDepth) is not PdfDictionary graphicsState ||
                    !graphicsState.Items.TryGetValue("SMask", out PdfObject? softMaskObject) ||
                    ResolveObject(
                        objects,
                        softMaskObject,
                        graphicsStateDepth + 1,
                        maximumObjectDepth,
                        out int softMaskDepth) is not PdfDictionary softMask ||
                    !softMask.Items.TryGetValue("G", out PdfObject? groupObject) ||
                    ResolveObject(
                        objects,
                        groupObject,
                        softMaskDepth + 1,
                        maximumObjectDepth,
                        out int groupDepth) is not PdfStream group) continue;
                AddInvokedFormContext(
                    group,
                    groupDepth,
                    resources,
                    aliases,
                    objects,
                    streams,
                    visitedContexts,
                    pending,
                    maximumObjectDepth);
            }
        }
    }

    private static void CollectResourceType3Streams(
        PdfDictionary resources,
        ColorSpaceAliases inheritedAliases,
        Dictionary<int, PdfIndirectObject> objects,
        List<ContentStreamContext> streams,
        Stack<(PdfDictionary Resources, ColorSpaceAliases Aliases, int Depth)> pending,
        int depth,
        int maximumObjectDepth) {
        if (!resources.Items.TryGetValue("Font", out PdfObject? fontsObject) ||
            ResolveObject(
                objects,
                fontsObject,
                depth + 1,
                maximumObjectDepth,
                out int fontsDepth) is not PdfDictionary fonts) return;
        foreach (PdfObject fontObject in fonts.Items.Values) {
            if (ResolveObject(
                    objects,
                    fontObject,
                    fontsDepth + 1,
                    maximumObjectDepth,
                    out int fontDepth) is not PdfDictionary font ||
                !string.Equals(
                    ResolveName(
                        font.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null,
                        objects,
                        maximumObjectDepth),
                    "Type3",
                    StringComparison.Ordinal) ||
                !font.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) ||
                ResolveObject(
                    objects,
                    charProcsObject,
                    fontDepth + 1,
                    maximumObjectDepth) is not PdfDictionary charProcs) continue;

            PdfDictionary? fontResources = ResolveResourcesDictionary(
                font,
                objects,
                inheritResources: false,
                maximumObjectDepth);
            ColorSpaceAliases aliases = fontResources == null
                ? inheritedAliases
                : CreateColorSpaceAliases(fontResources, objects, maximumObjectDepth);
            foreach (PdfObject charProc in charProcs.Items.Values) {
                CollectStreams(charProc, objects, aliases, streams, maximumObjectDepth);
            }
            if (fontResources != null) pending.Push((fontResources, aliases, fontDepth + 1));
        }
    }

    private static void CollectResourcePatternStreams(
        PdfDictionary resources,
        ColorSpaceAliases inheritedAliases,
        Dictionary<int, PdfIndirectObject> objects,
        List<ContentStreamContext> streams,
        List<ContentStreamContext> visitedContexts,
        Stack<(PdfDictionary Resources, ColorSpaceAliases Aliases, int Depth)> pending,
        int depth,
        int maximumObjectDepth) {
        if (!resources.Items.TryGetValue("Pattern", out PdfObject? patternsObject) ||
            ResolveObject(
                objects,
                patternsObject,
                depth + 1,
                maximumObjectDepth,
                out int patternsDepth) is not PdfDictionary patterns) return;
        foreach (PdfObject patternObject in patterns.Items.Values) {
            if (ResolveObject(
                    objects,
                    patternObject,
                    patternsDepth + 1,
                    maximumObjectDepth,
                    out int patternDepth) is not PdfStream pattern ||
                !pattern.Dictionary.Items.TryGetValue("PatternType", out PdfObject? patternType) ||
                ResolveObject(objects, patternType, patternDepth + 1, maximumObjectDepth) is not PdfNumber patternTypeNumber ||
                patternTypeNumber.Value != 1D) continue;

            ColorSpaceAliases aliases = inheritedAliases;
            PdfDictionary childResources = resources;
            if (pattern.Dictionary.Items.TryGetValue("Resources", out PdfObject? patternResourcesObject) &&
                ResolveObject(
                    objects,
                    patternResourcesObject,
                    patternDepth + 1,
                    maximumObjectDepth,
                    out int patternResourcesDepth) is PdfDictionary patternResources) {
                aliases = CreateColorSpaceAliases(patternResources, objects, maximumObjectDepth);
                childResources = patternResources;
                patternDepth = patternResourcesDepth;
            }
            if (ContainsContentStreamContext(visitedContexts, pattern, aliases)) continue;
            visitedContexts.Add(new ContentStreamContext(pattern, aliases));
            AddContentStream(pattern, aliases, streams);
            pending.Push((childResources, aliases, patternDepth + 1));
        }
    }

    private static void CollectResourceShadingContexts(
        PdfDictionary resources,
        ColorSpaceAliases aliases,
        Dictionary<int, PdfIndirectObject> objects,
        List<ShadingContext> shadings,
        int depth,
        int maximumObjectDepth) {
        if (resources.Items.TryGetValue("Shading", out PdfObject? shadingResourcesObject) &&
            ResolveObject(
                objects,
                shadingResourcesObject,
                depth + 1,
                maximumObjectDepth,
                out int shadingResourcesDepth) is PdfDictionary shadingResources) {
            foreach (PdfObject shadingObject in shadingResources.Items.Values) {
                AddResolvedShadingContext(
                    shadingObject,
                    shadingResourcesDepth + 1,
                    aliases,
                    objects,
                    shadings,
                    maximumObjectDepth);
            }
        }

        if (!resources.Items.TryGetValue("Pattern", out PdfObject? patternResourcesObject) ||
            ResolveObject(
                objects,
                patternResourcesObject,
                depth + 1,
                maximumObjectDepth,
                out int patternResourcesDepth) is not PdfDictionary patternResources) return;
        foreach (PdfObject patternObject in patternResources.Items.Values) {
            PdfObject? resolvedPattern = ResolveObject(
                objects,
                patternObject,
                patternResourcesDepth + 1,
                maximumObjectDepth,
                out int patternDepth);
            PdfDictionary? pattern = resolvedPattern switch {
                PdfDictionary dictionary => dictionary,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (pattern == null || !pattern.Items.TryGetValue("Shading", out PdfObject? shadingObject)) continue;
            AddResolvedShadingContext(
                shadingObject,
                patternDepth + 1,
                aliases,
                objects,
                shadings,
                maximumObjectDepth);
        }
    }

    private static void AddResolvedShadingContext(
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
        if (shading == null) return;
        for (int index = 0; index < shadings.Count; index++) {
            ShadingContext existing = shadings[index];
            if (ReferenceEquals(existing.Dictionary, shading) && existing.Aliases.SetEquals(aliases)) return;
        }
        shadings.Add(new ShadingContext(shading, aliases));
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

    private static void AddInvokedFormContext(
        PdfStream form,
        int formDepth,
        PdfDictionary inheritedResources,
        ColorSpaceAliases inheritedAliases,
        Dictionary<int, PdfIndirectObject> objects,
        List<ContentStreamContext> streams,
        List<ContentStreamContext> visitedContexts,
        Stack<(PdfDictionary Resources, ColorSpaceAliases Aliases, int Depth)> pending,
        int maximumObjectDepth) {
        if (!string.Equals(
                ResolveName(
                    form.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null,
                    objects,
                    maximumObjectDepth),
                "Form",
                StringComparison.Ordinal)) return;

        ColorSpaceAliases formAliases = inheritedAliases;
        PdfDictionary childResources = inheritedResources;
        if (form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourcesObject) &&
            ResolveObject(
                objects,
                formResourcesObject,
                formDepth + 1,
                maximumObjectDepth,
                out int formResourcesDepth) is PdfDictionary formResources) {
            formAliases = CreateColorSpaceAliases(formResources, objects, maximumObjectDepth);
            childResources = formResources;
            formDepth = formResourcesDepth;
        }

        if (ContainsContentStreamContext(visitedContexts, form, formAliases)) return;
        visitedContexts.Add(new ContentStreamContext(form, formAliases));
        AddContentStream(form, formAliases, streams);
        pending.Push((childResources, formAliases, formDepth + 1));
    }
}
