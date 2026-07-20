namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal IReadOnlyList<PdfRenderCapabilityDiagnostic> GetRenderCapabilityDiagnostics() {
        var diagnostics = new List<PdfRenderCapabilityDiagnostic>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var activeForms = new HashSet<PdfStream>();
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        CollectRenderCapabilityDiagnostics(GetContentStreamContent(), resources, diagnostics, seen, activeForms, 0);
        CollectAnnotationCapabilityDiagnostics(diagnostics, seen);
        return diagnostics.Count == 0 ? Array.Empty<PdfRenderCapabilityDiagnostic>() : diagnostics.AsReadOnly();
    }

    private void CollectRenderCapabilityDiagnostics(
        string content,
        PdfDictionary? resources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        int depth) {
        EnsureContentNestingBudget(depth);
        HashSet<string> unsupportedColorSpaces = GetUnsupportedColorSpaceResourceNames(resources);
        var invokedXObjects = new HashSet<string>(StringComparer.Ordinal);
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            string? capabilityId = GetOperatorCapabilityId(operation.Name);
            if (capabilityId != null) AddRenderDiagnostic(diagnostics, seen, capabilityId, operation.Name);
            if (operation.Name == "Do" &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string xObjectName) {
                invokedXObjects.Add(xObjectName);
            }
            if ((operation.Name == "cs" || operation.Name == "CS") &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string colorSpaceName &&
                unsupportedColorSpaces.Contains(colorSpaceName)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, colorSpaceName);
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
        CollectFontCapabilityDiagnostics(resources, diagnostics, seen);
        CollectPatternCapabilityDiagnostics(resources, diagnostics, seen);
        CollectGraphicsStateCapabilityDiagnostics(resources, diagnostics, seen);
        CollectXObjectCapabilityDiagnostics(resources, invokedXObjects, diagnostics, seen, activeForms, depth);
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

    private void CollectFontCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        foreach (PdfFontResource font in ResourceResolver.GetFontsForResources(resources, _objects).Values) {
            if (font.EmbeddedTrueTypeFont == null) AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.FontSubstitutionId, font.ResourceName);
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
                CollectRenderCapabilityDiagnostics(PdfEncoding.Latin1GetString(DecodeIfNeeded(stream)), formResources, diagnostics, seen, activeForms, depth + 1);
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

    private void CollectAnnotationCapabilityDiagnostics(List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        PdfArray? annotations = ResolveArray(_pageDict.Items.TryGetValue("Annots", out PdfObject? value) ? value : null);
        if (annotations == null) return;
        EnsureAnnotationBudget(annotations);
        for (int i = 0; i < annotations.Items.Count; i++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[i]);
            if (annotation == null || IsHiddenAnnotation(annotation) || TryGetNormalAppearanceStream(annotation, out _)) continue;
            string subtype = annotation.Get<PdfName>("Subtype")?.Name ?? "unknown";
            string capabilityId = PdfAnnotationFlattener.TryCreateSyntheticAppearanceStream(_objects, annotation, out _)
                ? PdfRenderCapabilities.SynthesizedAnnotationAppearanceId
                : PdfRenderCapabilities.AnnotationAppearanceId;
            AddRenderDiagnostic(diagnostics, seen, capabilityId, subtype + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
        }
    }

    private static void AddRenderDiagnostic(List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen, string capabilityId, string subject) {
        string key = capabilityId + "\n" + subject;
        if (seen.Add(key)) diagnostics.Add(new PdfRenderCapabilityDiagnostic(PdfRenderCapabilities.Get(capabilityId), subject));
    }
}
