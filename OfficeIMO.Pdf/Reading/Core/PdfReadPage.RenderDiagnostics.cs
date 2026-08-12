namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal IReadOnlyList<PdfRenderCapabilityDiagnostic> GetRenderCapabilityDiagnostics() {
        var diagnostics = new List<PdfRenderCapabilityDiagnostic>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        var type3GlyphBudget = new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        CollectRenderCapabilityDiagnostics(GetContentStreamContent(pageContentBudget), resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, 0);
        CollectAnnotationCapabilityDiagnostics(diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget);
        return diagnostics.Count == 0 ? Array.Empty<PdfRenderCapabilityDiagnostic>() : diagnostics.AsReadOnly();
    }

    private void CollectRenderCapabilityDiagnostics(
        string content,
        PdfDictionary? resources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        EnsureContentNestingBudget(depth);
        HashSet<string> unsupportedColorSpaces = GetUnsupportedColorSpaceResourceNames(resources);
        HashSet<string> approximatedIccColorSpaces = GetApproximatedIccColorSpaceResourceNames(resources);
        var invokedXObjects = new HashSet<string>(StringComparer.Ordinal);
        var invokedFonts = new HashSet<string>(StringComparer.Ordinal);
        var invokedShadings = new HashSet<string>(StringComparer.Ordinal);
        var invokedPatterns = new HashSet<string>(StringComparer.Ordinal);
        var invokedSoftMasks = new HashSet<PdfStream>();
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            string? capabilityId = GetOperatorCapabilityId(operation.Name);
            if (capabilityId != null) AddRenderDiagnostic(diagnostics, seen, capabilityId, operation.Name);
            if (operation.Name == "Do" &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string xObjectName) {
                invokedXObjects.Add(xObjectName);
            }
            if (operation.Name == "sh" && operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string shadingName) {
                invokedShadings.Add(shadingName);
            }
            if ((operation.Name == "cs" || operation.Name == "CS") &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string colorSpaceName) {
                if (unsupportedColorSpaces.Contains(colorSpaceName)) {
                    AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, colorSpaceName);
                } else if (approximatedIccColorSpaces.Contains(colorSpaceName)) {
                    AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, colorSpaceName);
                }
            }
            if (operation.InlineImage is PdfContentInlineImage inlineImage) {
                CollectImageColorSpaceCapabilityDiagnostic(
                    inlineImage.Dictionary,
                    resources,
                    diagnostics,
                    seen,
                    "inline-image");
            }
        },
        maxNestingDepth: _limits.MaxContentNestingDepth,
        maxOperands: _limits.MaxContentOperands);

        if (resources == null) return;
        HashSet<string> failedType3Fonts = CollectType3FontFailures(content, resources, pageContentBudget, type3GlyphBudget, invokedFonts, invokedPatterns, invokedSoftMasks, depth);
        CollectFontCapabilityDiagnostics(resources, invokedFonts, failedType3Fonts, diagnostics, seen);
        CollectShadingCapabilityDiagnostics(resources, invokedShadings, invokedPatterns, diagnostics, seen);
        CollectPatternCapabilityDiagnostics(resources, diagnostics, seen);
        CollectGraphicsStateCapabilityDiagnostics(resources, diagnostics, seen);
        CollectXObjectCapabilityDiagnostics(resources, invokedXObjects, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
        CollectAuxiliarySurfaceCapabilityDiagnostics(resources, invokedPatterns, invokedSoftMasks, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
    }

    private static string? GetOperatorCapabilityId(string op) {
        switch (op) {
            case "M": return PdfRenderCapabilities.MiterLimitId;
            case "ri": return PdfRenderCapabilities.RenderingIntentId;
            case "i": return PdfRenderCapabilities.FlatnessId;
            case "MP":
            case "DP": return PdfRenderCapabilities.MarkedPointId;
            case "d0":
            case "d1": return PdfRenderCapabilities.Type3MetricsId;
            default: return PdfContentOperators.IsStandard(op) ? null : PdfRenderCapabilities.UnknownOperatorId;
        }
    }

    private HashSet<string> CollectType3FontFailures(
        string content,
        PdfDictionary resources,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        HashSet<string> invokedFonts,
        HashSet<string> invokedPatterns,
        HashSet<PdfStream> invokedSoftMasks,
        int contentNestingDepth) {
        var failures = new HashSet<string>(StringComparer.Ordinal);
        var activeStreams = new HashSet<PdfStream>();
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        _ = PdfPageXObjectInvocationParser.Parse(
            content,
            Matrix2D.Identity,
            GetPageSize().Height,
            GetGraphicsStateResources(resources),
            GetColorSpaceResources(resources),
            GetOptionalContentVisibility(resources),
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            fonts: fonts,
            fontWidthProviders: ResourceResolver.GetFontWidthProvidersForResources(resources, _objects),
            type3TextVisitor: invocation => {
                bool supported = true;
                for (int i = 0; i < invocation.Glyphs.Count; i++) {
                    PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
                    if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
                        !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream stream) ||
                        !CanProjectType3VectorProgram(
                            stream,
                            type3.Resources,
                            Matrix2D.Multiply(glyph.Transform, type3.FontMatrix),
                            pageContentBudget,
                            type3GlyphBudget,
                            activeStreams,
                            contentNestingDepth + 1)) {
                        failures.Add(glyph.Font.ResourceName);
                        supported = false;
                    }
                }
                return supported;
            },
            type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
            visibleFontVisitor: fontName => {
                if (!string.IsNullOrEmpty(fontName)) invokedFonts.Add(fontName);
            },
            patternInvocationVisitor: patternName => invokedPatterns.Add(patternName),
            graphicsStateVisitor: state => {
                if (state.SoftMask?.Group is PdfStream group) invokedSoftMasks.Add(group);
            });
        return failures;
    }

    private bool CanProjectType3VectorProgram(
        PdfStream stream,
        PdfDictionary resources,
        Matrix2D programTransform,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        HashSet<PdfStream> activeStreams,
        int depth) {
        EnsureContentNestingBudget(depth);
        if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0 ||
            !activeStreams.Add(stream)) return false;
        try {
            string content;
            try {
                content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            } catch (IOException exception) when (exception is not PdfReadLimitException) {
                return false;
            }

            bool supported = true;
            Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
            foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                         content,
                         programTransform,
                         GetPageSize().Height,
                         GetGraphicsStateResources(resources),
                         GetColorSpaceResources(resources),
                         GetOptionalContentVisibility(resources),
                         maxOperations: _limits.MaxContentOperations,
                         maxNestingDepth: _limits.MaxContentNestingDepth,
                         maxOperands: _limits.MaxContentOperands,
                         fonts: fonts,
                         fontWidthProviders: ResourceResolver.GetFontWidthProvidersForResources(resources, _objects),
                         type3TextVisitor: nested => {
                             for (int index = 0; index < nested.Glyphs.Count; index++) {
                                 PdfPageType3GlyphInvocation glyph = nested.Glyphs[index];
                                 if (glyph.Font.Type3 is not PdfType3FontResource nestedType3 ||
                                     !nestedType3.TryGetGlyph(glyph.CharacterCode, out PdfStream nestedStream) ||
                                     !CanProjectType3VectorProgram(
                                         nestedStream,
                                         nestedType3.Resources,
                                         Matrix2D.Multiply(glyph.Transform, nestedType3.FontMatrix),
                                         pageContentBudget,
                                         type3GlyphBudget,
                                         activeStreams,
                                         depth + 1)) {
                                     supported = false;
                                 }
                             }
                             return supported;
                         },
                         type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
                         unsupportedTextVisitor: () => supported = false,
                         unsupportedGraphicsEffectVisitor: () => supported = false,
                         unsupportedPatternVisitor: () => supported = false,
                         unsupportedColorVisitor: () => supported = false)) {
                if (invocation.InlineImage != null ||
                    !TryGetFormStream(resources, invocation.Name, out PdfStream form)) {
                    supported = false;
                    continue;
                }
                if (HasUnsupportedType3FormGroup(form.Dictionary)) {
                    supported = false;
                    continue;
                }
                PdfDictionary formResources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? value) ? value : null) ?? resources;
                if (!CanProjectType3VectorProgram(
                        form,
                        formResources,
                        ApplyFormMatrix(invocation.Transform, form.Dictionary),
                        pageContentBudget,
                        type3GlyphBudget,
                        activeStreams,
                        depth + 1)) supported = false;
            }
            return supported;
        } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
            return false;
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private void CollectFontCapabilityDiagnostics(PdfDictionary resources, HashSet<string> invokedFonts, HashSet<string> failedType3Fonts, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        foreach (PdfFontResource font in ResourceResolver.GetFontsForResources(resources, _objects).Values) {
            if (!invokedFonts.Contains(font.ResourceName) || font.EmbeddedTrueTypeFont != null) continue;
            string capabilityId;
            if (string.Equals(font.FontSubtype, "Type3", StringComparison.Ordinal)) {
                if (font.Type3 != null && !failedType3Fonts.Contains(font.ResourceName)) continue;
                capabilityId = PdfRenderCapabilities.Type3FontSubstitutionId;
            } else if (font.EmbeddedProgramSubtype is "Type1C" or "CIDFontType0C" or "CFF") {
                capabilityId = PdfRenderCapabilities.CffFontSubstitutionId;
            } else {
                capabilityId = PdfRenderCapabilities.FontSubstitutionId;
            }
            AddRenderDiagnostic(diagnostics, seen, capabilityId, font.ResourceName);
        }
    }

    private HashSet<string> GetUnsupportedColorSpaceResourceNames(PdfDictionary? resources) {
        var unsupported = new HashSet<string>(StringComparer.Ordinal);
        if (resources == null) return unsupported;
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return unsupported;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (!TryReadColorSpaceResource(entry.Value, out _)) {
                unsupported.Add(entry.Key);
            }
        }

        return unsupported;
    }

    private HashSet<string> GetApproximatedIccColorSpaceResourceNames(PdfDictionary? resources) {
        var approximated = new HashSet<string>(StringComparer.Ordinal);
        if (resources == null) return approximated;
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return approximated;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (TryReadColorSpaceResource(entry.Value, out PdfPageColorSpace colorSpace) && colorSpace.UsesIccApproximation) {
                approximated.Add(entry.Key);
            }
        }
        return approximated;
    }

    private void CollectShadingCapabilityDiagnostics(
        PdfDictionary resources,
        IReadOnlyCollection<string> invokedShadings,
        IReadOnlyCollection<string> invokedPatterns,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen) {
        PdfDictionary? shadings = ResolveDictionary(resources.Items.TryGetValue("Shading", out PdfObject? shadingValue) ? shadingValue : null);
        foreach (string name in invokedShadings) {
            if (shadings?.Items.TryGetValue(name, out PdfObject? shading) == true) {
                CollectOneShadingCapabilityDiagnostic(shading, name, diagnostics, seen);
            } else {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, name);
            }
        }

        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? patternValue) ? patternValue : null);
        foreach (string name in invokedPatterns) {
            if (patterns?.Items.TryGetValue(name, out PdfObject? patternValueObject) != true) continue;
            PdfDictionary? pattern = ResolveDictionary(patternValueObject);
            if (TryReadInteger(pattern?.Items.TryGetValue("PatternType", out PdfObject? typeValue) == true ? typeValue : null) != 2) continue;
            if (pattern?.Items.TryGetValue("Shading", out PdfObject? shading) == true) {
                CollectOneShadingCapabilityDiagnostic(shading, name, diagnostics, seen);
            } else {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, name);
            }
        }
    }

    private void CollectOneShadingCapabilityDiagnostic(
        PdfObject? value,
        string subject,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen) {
        PdfDictionary? shading = ResolveDictionary(value);
        if (shading == null || !shading.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ||
            !TryReadColorSpaceResource(colorSpaceObject, out PdfPageColorSpace colorSpace)) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, subject);
        } else if (colorSpace.UsesIccApproximation) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, subject);
        }
        if (!TryReadShading(value, out _)) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, subject);
        }
    }

    private void CollectPatternCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? value) ? value : null);
        if (patterns == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in patterns.Items) {
            PdfObject? resolved = ResolveObject(entry.Value);
            PdfDictionary? pattern = resolved switch {
                PdfDictionary dictionary => dictionary,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (pattern?.Get<PdfNumber>("PatternType")?.Value == 1D) {
                if (!IsStructurallySupportedTilingPattern(resolved, pattern)) {
                    AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedTilingPatternId, entry.Key);
                }
            }
        }
    }

    private bool IsStructurallySupportedTilingPattern(PdfObject? resolved, PdfDictionary pattern) {
        int? paintType = TryReadInteger(pattern.Items.TryGetValue("PaintType", out PdfObject? paintTypeObject) ? paintTypeObject : null);
        int? tilingType = TryReadInteger(pattern.Items.TryGetValue("TilingType", out PdfObject? tilingTypeObject) ? tilingTypeObject : null);
        Matrix2D matrix = pattern.Items.TryGetValue("Matrix", out PdfObject? matrixObject)
            ? ReadPatternMatrix(matrixObject)
            : Matrix2D.Identity;
        return resolved is PdfStream &&
            (paintType == 1 || paintType == 2) &&
            tilingType >= 1 && tilingType <= 3 &&
            TryReadRectangle(pattern.Items.TryGetValue("BBox", out PdfObject? boxObject) ? boxObject : null, out (double X1, double Y1, double X2, double Y2) box) &&
            box.X2 > box.X1 && box.Y2 > box.Y1 &&
            ResolveObject(pattern.Items.TryGetValue("XStep", out PdfObject? xStepObject) ? xStepObject : null) is PdfNumber xStep &&
            ResolveObject(pattern.Items.TryGetValue("YStep", out PdfObject? yStepObject) ? yStepObject : null) is PdfNumber yStep &&
            !double.IsNaN(xStep.Value) && !double.IsInfinity(xStep.Value) && Math.Abs(xStep.Value) > 0.0000001D &&
            !double.IsNaN(yStep.Value) && !double.IsInfinity(yStep.Value) && Math.Abs(yStep.Value) > 0.0000001D &&
            IsUsableTilingPatternMatrix(matrix);
    }

    private void CollectAuxiliarySurfaceCapabilityDiagnostics(
        PdfDictionary resources,
        IReadOnlyCollection<string> invokedPatterns,
        IReadOnlyCollection<PdfStream> invokedSoftMasks,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? patternObject) ? patternObject : null);
        foreach (string patternName in invokedPatterns) {
            if (patterns?.Items.TryGetValue(patternName, out PdfObject? patternValue) != true ||
                ResolveObject(patternValue) is not PdfStream patternStream ||
                TryReadInteger(patternStream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? typeValue) ? typeValue : null) != 1) continue;
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(patternStream, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
        }

        foreach (PdfStream softMaskGroup in invokedSoftMasks) {
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(softMaskGroup, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
        }
    }

    private void CollectOneAuxiliarySurfaceCapabilityDiagnostics(
        PdfStream stream,
        PdfDictionary parentResources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        if (!activeForms.Add(stream)) return;
        try {
            PdfDictionary? resources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? parentResources;
            string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            CollectRenderCapabilityDiagnostics(content, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth + 1);
        } finally {
            activeForms.Remove(stream);
        }
    }

    private void CollectGraphicsStateCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        PdfDictionary? states = ResolveDictionary(resources.Items.TryGetValue("ExtGState", out PdfObject? value) ? value : null);
        if (states == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in states.Items) {
            PdfDictionary? state = ResolveDictionary(entry.Value);
            if (state == null) continue;
            if (state.Items.TryGetValue("BM", out _) && !ReadBlendMode(state).HasValue) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedBlendModeId, entry.Key);
            }
            if (state.Items.TryGetValue("SMask", out PdfObject? mask) &&
                ResolveObject(mask) is not PdfName { Name: "None" } &&
                ReadSoftMask(state) == null) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedSoftMaskId, entry.Key);
            }
        }
    }

    private void CollectXObjectCapabilityDiagnostics(
        PdfDictionary resources,
        HashSet<string> invokedXObjects,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        PdfDictionary? xObjects = ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? value) ? value : null);
        if (xObjects == null) return;
        foreach (string invokedName in invokedXObjects) {
            if (!xObjects.Items.TryGetValue(invokedName, out PdfObject? xObject)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, invokedName);
                continue;
            }

            var entry = new KeyValuePair<string, PdfObject>(invokedName, xObject);
            if (ResolveObject(entry.Value) is not PdfStream stream) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, entry.Key);
                continue;
            }

            string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                CollectImageColorSpaceCapabilityDiagnostic(
                    stream.Dictionary,
                    resources,
                    diagnostics,
                    seen,
                    entry.Key);
                if (RequiresOptionalImageCodec(stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? filterObject) ? filterObject : null)) AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.OptionalImageCodecId, entry.Key);
                continue;
            }
            if (!string.Equals(subtype, "Form", StringComparison.Ordinal)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, entry.Key + ":" + (subtype ?? "unknown"));
                continue;
            }

            if (!activeForms.Add(stream)) continue;
            try {
                PdfDictionary? formResources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceObject) ? formResourceObject : null) ?? resources;
                CollectRenderCapabilityDiagnostics(PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream)), formResources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth + 1);
            } finally {
                activeForms.Remove(stream);
            }
        }
    }

    private void CollectImageColorSpaceCapabilityDiagnostic(
        PdfDictionary image,
        PdfDictionary? resources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        string imageName) {
        if (!image.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)) {
            return;
        }

        if (ResourceResolver.CanProjectImageColorSpace(image, resources, _objects)) {
            PdfObject? diagnosticColorSpace = colorSpaceObject;
            if (ResolveObject(colorSpaceObject) is PdfName resourceName) {
                PdfDictionary? colorSpaces = ResolveDictionary(resources?.Items.TryGetValue("ColorSpace", out PdfObject? value) == true ? value : null);
                if (colorSpaces?.Items.TryGetValue(resourceName.Name, out PdfObject? resourceColorSpace) == true) diagnosticColorSpace = resourceColorSpace;
            }
            if (TryReadColorSpaceResource(diagnosticColorSpace, out PdfPageColorSpace projectedColorSpace) && projectedColorSpace.UsesIccApproximation) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, imageName);
            }
            return;
        }

        PdfObject? resolved = ResolveObject(colorSpaceObject);
        string subject = imageName;
        if (resolved is PdfName name) {
            subject = name.Name;
        }

        AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, subject);
    }

    private bool RequiresOptionalImageCodec(PdfObject? value) {
        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfName name) return name.Name is "JPXDecode";
        if (resolved is not PdfArray array) return false;
        for (int i = 0; i < array.Items.Count; i++) if (RequiresOptionalImageCodec(array.Items[i])) return true;
        return false;
    }

    private void CollectAnnotationCapabilityDiagnostics(
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        PdfArray? annotations = ResolveArray(_pageDict.Items.TryGetValue("Annots", out PdfObject? value) ? value : null);
        if (annotations == null) return;
        EnsureAnnotationBudget(annotations);
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        for (int i = 0; i < annotations.Items.Count; i++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[i]);
            if (annotation == null || IsHiddenAnnotation(annotation) || HasNoVisibleAnnotationArea(annotation)) continue;
            string subtype = annotation.Get<PdfName>("Subtype")?.Name ?? "unknown";
            if (!TryGetRenderableAnnotationAppearanceStream(annotation, out PdfStream appearance, out bool synthesized)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.AnnotationAppearanceId, subtype + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
                continue;
            }
            if (synthesized) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.SynthesizedAnnotationAppearanceId, subtype + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
            }
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(appearance, pageResources ?? new PdfDictionary(), diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, 0);
        }
    }

    private bool HasNoVisibleAnnotationArea(PdfDictionary annotation) {
        PdfArray? rectangle = ResolveArray(annotation.Items.TryGetValue("Rect", out PdfObject? value) ? value : null);
        if (rectangle == null || rectangle.Items.Count < 4 ||
            ResolveObject(rectangle.Items[0]) is not PdfNumber x1 ||
            ResolveObject(rectangle.Items[1]) is not PdfNumber y1 ||
            ResolveObject(rectangle.Items[2]) is not PdfNumber x2 ||
            ResolveObject(rectangle.Items[3]) is not PdfNumber y2) {
            return false;
        }
        return x1.Value == x2.Value || y1.Value == y2.Value;
    }

    private static void AddRenderDiagnostic(List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen, string capabilityId, string subject) {
        string key = capabilityId + "\n" + subject;
        if (seen.Add(key)) diagnostics.Add(new PdfRenderCapabilityDiagnostic(PdfRenderCapabilities.Get(capabilityId), subject));
    }
}
