namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool HasMalformedStrictInvocation(
        string content,
        PdfDictionary? resources,
        PageContentBudget pageContentBudget,
        HashSet<PdfStream> activeStreams,
        int contentNestingDepth,
        Dictionary<(PdfStream Stream, PdfDictionary? Resources), bool>? resultCache = null) {
        resultCache ??= new Dictionary<(PdfStream Stream, PdfDictionary? Resources), bool>();
        bool malformed = false;
        string? fontName = null;
        var fontStack = new Stack<string?>();
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        PdfContentStreamInterpreter.InterpretUntil(
            content,
            _limits.MaxContentOperations,
            operation => {
                switch (operation.Name) {
                    case "q":
                        fontStack.Push(fontName);
                        break;
                    case "Q":
                        fontName = fontStack.Count > 0 ? fontStack.Pop() : null;
                        break;
                    case "Tf" when operation.Operands.Count == 2 && operation.Operands[0] is string selectedFont:
                        fontName = selectedFont;
                        break;
                    case "cs": case "CS": case "sh":
                        malformed = operation.HasInvalidOperands ||
                            operation.Operands.Count != 1 ||
                            operation.Operands[0] is not string;
                        break;
                    case "Do":
                        if (operation.HasInvalidOperands || operation.Operands.Count != 1 || operation.Operands[0] is not string xObjectName) {
                            malformed = true;
                            break;
                        }
                        if (TryResolvePatternForm(resources, xObjectName, out PdfStream form)) {
                            PdfDictionary? formResources = ResolveDictionary(
                                form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourcesObject)
                                    ? formResourcesObject
                                    : null) ?? resources;
                            malformed = NestedPatternStreamHasMalformedStrictInvocation(
                                form,
                                formResources,
                                pageContentBudget,
                                activeStreams,
                                contentNestingDepth,
                                resultCache);
                        }
                        break;
                    case "Tj": case "TJ": case "'": case "\"":
                        if (fontName is string activeFontName &&
                            fonts.TryGetValue(activeFontName, out PdfFontResource? font) &&
                            font.Type3 is PdfType3FontResource type3) {
                            foreach (byte[] bytes in GetShownTextBytes(operation)) {
                                for (int index = 0; index < bytes.Length && !malformed; index++) {
                                    if (!type3.TryGetGlyph(bytes[index], out PdfStream glyph)) continue;
                                    malformed = NestedPatternStreamHasMalformedStrictInvocation(
                                        glyph,
                                        type3.Resources,
                                        pageContentBudget,
                                        activeStreams,
                                        contentNestingDepth,
                                        resultCache);
                                }
                                if (malformed) break;
                            }
                        }
                        break;
                }
                return !malformed;
            },
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            dispatchInvalidOperations: true);
        return malformed;
    }

    private bool NestedPatternStreamHasMalformedStrictInvocation(
        PdfStream stream,
        PdfDictionary? resources,
        PageContentBudget pageContentBudget,
        HashSet<PdfStream> activeStreams,
        int contentNestingDepth,
        Dictionary<(PdfStream Stream, PdfDictionary? Resources), bool> resultCache) {
        var cacheKey = (stream, resources);
        if (resultCache.TryGetValue(cacheKey, out bool cached)) return cached;
        if (!activeStreams.Add(stream)) return false;
        try {
            EnsureContentNestingBudget(contentNestingDepth + 1);
            string nestedContent = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            bool malformed = HasMalformedStrictInvocation(
                nestedContent,
                resources,
                pageContentBudget,
                activeStreams,
                contentNestingDepth + 1,
                resultCache);
            resultCache[cacheKey] = malformed;
            return malformed;
        } finally {
            activeStreams.Remove(stream);
        }
    }
}
