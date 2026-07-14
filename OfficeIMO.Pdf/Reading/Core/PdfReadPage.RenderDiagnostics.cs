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
        foreach (string op in PdfRenderOperatorScanner.ReadOperators(content, _limits.MaxContentOperations)) {
            string? capabilityId = GetOperatorCapabilityId(op);
            if (capabilityId != null) AddRenderDiagnostic(diagnostics, seen, capabilityId, op);
        }

        if (resources == null) return;
        CollectFontCapabilityDiagnostics(resources, diagnostics, seen);
        CollectColorSpaceCapabilityDiagnostics(resources, diagnostics, seen);
        CollectPatternCapabilityDiagnostics(resources, diagnostics, seen);
        CollectGraphicsStateCapabilityDiagnostics(resources, diagnostics, seen);
        CollectXObjectCapabilityDiagnostics(resources, diagnostics, seen, activeForms, depth);
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
            default: return PdfRenderOperatorScanner.IsKnownManagedOperator(op) ? null : PdfRenderCapabilities.UnknownOperatorId;
        }
    }

    private void CollectFontCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        foreach (PdfFontResource font in ResourceResolver.GetFontsForResources(resources, _objects).Values) {
            if (font.EmbeddedTrueTypeFont == null) AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.FontSubstitutionId, font.ResourceName);
        }
    }

    private void CollectColorSpaceCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (!TryReadColorSpaceResource(entry.Value, out _)) AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, entry.Key);
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
            if (pattern?.Get<PdfNumber>("PatternType")?.Value == 1D && !IsStructurallySupportedTilingPattern(resolved, pattern)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedTilingPatternId, entry.Key);
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
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        int depth) {
        PdfDictionary? xObjects = ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? value) ? value : null);
        if (xObjects == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
            if (ResolveObject(entry.Value) is not PdfStream stream) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, entry.Key);
                continue;
            }

            string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
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
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.AnnotationAppearanceId, subtype + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
        }
    }

    private static void AddRenderDiagnostic(List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen, string capabilityId, string subject) {
        string key = capabilityId + "\n" + subject;
        if (seen.Add(key)) diagnostics.Add(new PdfRenderCapabilityDiagnostic(PdfRenderCapabilities.Get(capabilityId), subject));
    }
}

internal static class PdfRenderOperatorScanner {
    private static readonly HashSet<string> KnownOperators = new HashSet<string>(StringComparer.Ordinal) {
        "q", "Q", "cm", "w", "J", "j", "d", "gs", "CS", "cs", "SC", "SCN", "sc", "scn", "G", "g", "RG", "rg", "K", "k",
        "m", "l", "c", "v", "y", "h", "re", "S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n", "W", "W*", "sh",
        "BT", "ET", "Tc", "Tw", "Tz", "TL", "Tf", "Tr", "Ts", "Td", "TD", "Tm", "T*", "Tj", "TJ", "'", "\"",
        "Do", "BI", "ID", "EI", "BMC", "BDC", "EMC", "BX", "EX", "M", "ri", "i", "MP", "DP", "d0", "d1"
    };

    public static bool IsKnownManagedOperator(string op) => KnownOperators.Contains(op);

    public static IReadOnlyList<string> ReadOperators(string content, int maxOperations) {
        if (string.IsNullOrEmpty(content)) return Array.Empty<string>();
        var result = new List<string>();
        int index = 0;
        while (index < content.Length) {
            SkipWhiteSpaceAndComments(content, ref index);
            if (index >= content.Length) break;
            char current = content[index];
            if (current == '/') { SkipName(content, ref index); continue; }
            if (current == '(') { SkipLiteralString(content, ref index); continue; }
            if (current == '<') { SkipAngleObject(content, ref index); continue; }
            if (current == '[') { SkipBalanced(content, ref index, '[', ']'); continue; }
            if (IsNumberStart(current)) { SkipToken(content, ref index); continue; }
            string token = ReadToken(content, ref index);
            if (token.Length == 0 || token == "true" || token == "false" || token == "null") continue;
            if (result.Count >= maxOperations) throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperations, maxOperations, result.Count + 1);
            result.Add(token);
            if (token == "BI") SkipInlineImage(content, ref index);
        }

        return result;
    }

    private static void SkipInlineImage(string content, ref int index) {
        int id = FindDelimitedToken(content, index, "ID");
        if (id < 0) { index = content.Length; return; }
        int dataStart = id + 2;
        if (dataStart < content.Length && char.IsWhiteSpace(content[dataStart])) dataStart++;
        int end = FindDelimitedToken(content, dataStart, "EI");
        index = end < 0 ? content.Length : end + 2;
    }

    private static int FindDelimitedToken(string content, int start, string token) {
        int index = start;
        while ((index = content.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            bool before = index == 0 || IsDelimiter(content[index - 1]);
            int afterIndex = index + token.Length;
            bool after = afterIndex >= content.Length || IsDelimiter(content[afterIndex]);
            if (before && after) return index;
            index++;
        }
        return -1;
    }

    private static void SkipWhiteSpaceAndComments(string content, ref int index) {
        while (index < content.Length) {
            if (char.IsWhiteSpace(content[index]) || content[index] == '\0') { index++; continue; }
            if (content[index] != '%') return;
            while (index < content.Length && content[index] != '\r' && content[index] != '\n') index++;
        }
    }

    private static void SkipName(string content, ref int index) { index++; SkipToken(content, ref index); }

    private static void SkipToken(string content, ref int index) {
        while (index < content.Length && !IsDelimiter(content[index])) index++;
    }

    private static string ReadToken(string content, ref int index) {
        int start = index;
        SkipToken(content, ref index);
        return index == start ? content[index++].ToString() : content.Substring(start, index - start);
    }

    private static void SkipLiteralString(string content, ref int index) {
        int depth = 0;
        while (index < content.Length) {
            char current = content[index++];
            if (current == '\\' && index < content.Length) { index++; continue; }
            if (current == '(') depth++;
            else if (current == ')' && --depth <= 0) return;
        }
    }

    private static void SkipAngleObject(string content, ref int index) {
        if (index + 1 < content.Length && content[index + 1] == '<') SkipBalanced(content, ref index, '<', '>');
        else {
            index++;
            while (index < content.Length && content[index++] != '>') { }
        }
    }

    private static void SkipBalanced(string content, ref int index, char open, char close) {
        int depth = 0;
        while (index < content.Length) {
            char current = content[index++];
            if (current == '(') { index--; SkipLiteralString(content, ref index); continue; }
            if (current == open) depth++;
            else if (current == close && --depth <= 0) return;
        }
    }

    private static bool IsNumberStart(char value) => value == '+' || value == '-' || value == '.' || (value >= '0' && value <= '9');
    private static bool IsDelimiter(char value) => char.IsWhiteSpace(value) || value == '\0' || value == '(' || value == ')' || value == '<' || value == '>' || value == '[' || value == ']' || value == '{' || value == '}' || value == '/' || value == '%';
}
