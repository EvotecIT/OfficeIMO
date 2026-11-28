
namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool TryRenderNativeParagraphFlowBreaks(
        INativePdfFlow pdf, WordParagraph paragraph, (int Level, string Marker)? marker,
        System.Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeNoteNumbering footnoteNumbersById, WordPdfExportOptions? options,
        IReadOnlyDictionary<OfficeXmlNode, string> headingDestinations,
        NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap,
        WordParagraph? nextParagraph) {
        // Each split clones the remaining suffix. Bound the number of visible boundaries
        // before the first clone so one authored paragraph cannot amplify that work.
        const int maximumFlowBreaks = 256;
        int flowBreaks = 0;
        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OfficeXmlNode boundary in run.ProjectedRunChildNodes()) {
                if (IsNativeFlowBreak(boundary) && ++flowBreaks > maximumFlowBreaks)
                    throw new NotSupportedException("A paragraph has more visible page or column breaks than PDF conversion supports.");
            }
        }
        if (!TrySplitNativeParagraphAtVisibleFlowBreak(paragraph, out string breakType, out WordParagraph? before, out WordParagraph? after)) return false;

        // Splitting can leave no renderable fragment. Diagnose the authored frame
        // once, then keep its synthetic fragments out of the diagnostic count.
        if (options != null) {
            RecordNativeParagraphFrameDiagnostics(paragraph, options, "body paragraph");
            RecordNativeParagraphBarTabDiagnostics(paragraph, options, "body paragraph",
                omitted: before == null && after == null);
        }
        if (HasNativePageBreakBefore(paragraph)) RenderNativePageBreakBefore(pdf, paragraph);
        OfficeXmlNode source = paragraph.Node();
        if (source.Exists && !GetNativeParagraphBookmarkNames(paragraph).Any() &&
            headingDestinations.TryGetValue(source, out string? generatedDestination)) pdf.Bookmark(generatedDestination);
        while (true) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            if (before != null) {
                OfficeXmlNode beforeNode = before.Node();
                if (beforeNode.Exists) {
                    options?.DiagnosedFrameParagraphs.Add(beforeNode);
                    options?.DiagnosedBarTabElements.Add(beforeNode);
                }
                before.PageBreakBeforeOverride = false;
                before.KeepWithNextOverride = false;
                before.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                RenderNativeParagraph(pdf, before, marker, getMarker, System.Array.Empty<int>(), footnoteNumbersById,
                    options, headingDestinations, nativeDefaults, nativeFontMap, null);
                marker = null;
            }
            if (breakType == "column") pdf.ColumnBreak();
            else pdf.PageBreak(preserveEmptyPage: true);
            // Word moves the paragraph mark after a column break into the next column,
            // where it takes an empty line; after a page break it stays on the previous page.
            if (after == null && (breakType == "column" || breakType == "page" &&
                !UsesModernNativeWordLayout(paragraph._document) &&
                paragraph._document.CompatibilitySettings.SplitPageBreakAndParagraphMark)) {
                OfficeXmlNode mark = CreateNativeDetachedBodyElement(paragraph._document, "p");
                OfficeXmlNode markProperties = source.WChild("pPr");
                if (markProperties.Exists) {
                    OfficeXmlNode copiedProperties = mark.Append(markProperties.Clone());
                    OfficeXmlNode sectionProperties = copiedProperties.WChild("sectPr");
                    if (sectionProperties.Exists) sectionProperties.Remove();
                }
                after = WordConverterNodes.ParagraphView(paragraph._document, mark);
            }
            if (after == null) return true;
            OfficeXmlNode afterNode = after.Node();
            if (afterNode.Exists) {
                options?.DiagnosedFrameParagraphs.Add(afterNode);
                options?.DiagnosedBarTabElements.Add(afterNode);
            }
            after.PageBreakBeforeOverride = false;
            after.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
            OfficeXmlNode properties = after.Node().GetOrPrependChild(WordXml.W, "pPr");
            // The continuation is the same paragraph, so it has no space before; an explicit
            // off also stops a style or document default from turning automatic spacing back on.
            OfficeXmlNode continuationSpacing = properties.WChild("spacing");
            if (continuationSpacing.Exists) continuationSpacing.SetAttribute(WordXml.W, "beforeAutospacing", "0");
            OfficeXmlNode indentation = properties.WChild("ind");
            if (indentation.IsNone) indentation = properties.InsertOrdered(properties.CreateElement(WordXml.W, "ind"), WPPrView.Order);
            indentation.SetAttribute(WordXml.W, "firstLine", "0");
            indentation.SetAttribute(WordXml.W, "firstLineChars", null);
            indentation.SetAttribute(WordXml.W, "hanging", null);
            indentation.SetAttribute(WordXml.W, "hangingChars", null);
            if (!TrySplitNativeParagraphAtVisibleFlowBreak(after, out breakType, out before, out WordParagraph? remaining)) {
                RenderNativeParagraph(pdf, after, marker, getMarker, System.Array.Empty<int>(), footnoteNumbersById,
                    options, headingDestinations, nativeDefaults, nativeFontMap, nextParagraph);
                return true;
            }
            after = remaining;
        }
    }

    /// <summary>Whether the element is a <c>w:br</c> that ends a page or a column.</summary>
    private static bool IsNativeFlowBreak(OfficeXmlNode element) {
        if (!element.IsW("br")) return false;
        string? type = element.WAttr("type");
        return type == "page" || type == "column";
    }

    /// <summary>
    /// A new, detached element of the main story, such as a <c>w:p</c> or <c>w:tbl</c>, made as the body's own child
    /// so it reads like any other element of that kind.
    /// </summary>
    private static OfficeXmlNode CreateNativeDetachedBodyElement(WordDocument document, string localName) {
        OfficeXmlNode body = document.BodyNode();
        if (body.IsNone) throw new InvalidOperationException("The document has no body to create content for.");
        return body.CreateElement(WordXml.W, localName);
    }

    private static bool TrySplitNativeParagraphAtVisibleFlowBreak(WordParagraph paragraph, out string breakType, out WordParagraph? before, out WordParagraph? after) {
        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            OfficeXmlNode boundary = run.ProjectedRunChildNodes().FirstOrDefault(IsNativeFlowBreak);
            if (boundary.Exists) {
                breakType = boundary.WAttr("type")!;
                return TrySplitNativeParagraphAtVisibleBreak(paragraph, breakType, out before, out after);
            }
        }
        breakType = "page"; before = null; after = null;
        return false;
    }

    private static bool TrySplitNativeParagraphAtVisibleBreak(WordParagraph paragraph, string breakType, out WordParagraph? before, out WordParagraph? after) {
        before = null;
        after = null;
        if (!paragraph.Node().DescendantElements(WordXml.W, "br").Any(wordBreak => wordBreak.WAttr("type") == breakType)) return false;
        var visibleBreaks = new HashSet<OfficeXmlNode>();
        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OfficeXmlNode wordBreak in run.ProjectedRunChildNodes()) {
                if (wordBreak.IsW("br") && wordBreak.WAttr("type") == breakType) visibleBreaks.Add(wordBreak);
            }
        }
        if (visibleBreaks.Count == 0) { before = null; after = null; return false; }
        return TrySplitNativeParagraphAtBreak(paragraph, breakType, visibleBreaks, out before, out after);
    }
    private static bool TrySplitNativeParagraphAtBreak(WordParagraph paragraph, string breakType,
        HashSet<OfficeXmlNode>? visibleBreaks, out WordParagraph? before, out WordParagraph? after) {
        before = null;
        after = null;
        OfficeXmlNode source = paragraph.Node();
        if (source.IsNone) {
            return false;
        }

        OfficeXmlNode beforeParagraph = CreateNativeDetachedBodyElement(paragraph._document, "p");
        OfficeXmlNode afterParagraph = CreateNativeDetachedBodyElement(paragraph._document, "p");
        OfficeXmlNode sourceProperties = source.WChild("pPr");
        if (sourceProperties.Exists) {
            beforeParagraph.Append(sourceProperties.Clone());
            afterParagraph.Append(sourceProperties.Clone());
        }

        bool sawBreak = false;
        foreach (OfficeXmlNode child in source.Elements()) {
            if (child.IsW("pPr")) {
                continue;
            }

            if (!sawBreak &&
                TrySplitNativeOpenXmlAtBreak(child, breakType, visibleBreaks, out OfficeXmlNode beforeChild, out OfficeXmlNode afterChild)) {
                if (beforeChild.Exists) {
                    beforeParagraph.Append(beforeChild);
                }

                if (afterChild.Exists) {
                    afterParagraph.Append(afterChild);
                }

                sawBreak = true;
                continue;
            }

            if (sawBreak) {
                afterParagraph.Append(child.Clone());
            } else {
                beforeParagraph.Append(child.Clone());
            }
        }

        if (!sawBreak) {
            return false;
        }

        OfficeXmlNode boundary = source.DescendantElements(WordXml.W, "br").First(wordBreak =>
            wordBreak.WAttr("type") == breakType && (visibleBreaks == null || visibleBreaks.Contains(wordBreak)));
        WordConverterNodes.PreserveFragmentContext(source, beforeParagraph, afterParagraph, boundary);

        if (HasNativeRenderableOpenXmlContent(beforeParagraph)) {
            before = WordConverterNodes.ParagraphView(paragraph._document, beforeParagraph, beforeParagraph.Child(WordXml.W, "bookmarkStart"));
        }

        if (HasNativeRenderableOpenXmlContent(afterParagraph)) {
            after = WordConverterNodes.ParagraphView(paragraph._document, afterParagraph, afterParagraph.Child(WordXml.W, "bookmarkStart"));
        }

        return true;
    }

    private static bool TrySplitNativeOpenXmlAtBreak(
        OfficeXmlNode element,
        string breakType,
        HashSet<OfficeXmlNode>? visibleBreaks,
        out OfficeXmlNode before,
        out OfficeXmlNode after) {
        before = default;
        after = default;
        if (element.IsW("br") && element.WAttr("type") == breakType &&
            (visibleBreaks == null || visibleBreaks.Contains(element))) {
            return true;
        }

        if (!element.HasChildren()) {
            return false;
        }

        OfficeXmlNode beforeElement = element.CloneShallow();
        OfficeXmlNode afterElement = element.CloneShallow();
        bool sawBreak = false;
        bool containsBreak = false;
        foreach (OfficeXmlNode child in element.Elements()) {
            // Run/control properties apply to both fragments of their container.
            if (child.IsW("rPr") || child.IsW("sdtPr") || child.IsW("sdtEndPr")) {
                beforeElement.Append(child.Clone());
                afterElement.Append(child.Clone());
                continue;
            }
            if (!sawBreak &&
                TrySplitNativeOpenXmlAtBreak(child, breakType, visibleBreaks, out OfficeXmlNode beforeChild, out OfficeXmlNode afterChild)) {
                containsBreak = true;
                if (beforeChild.Exists) {
                    beforeElement.Append(beforeChild);
                }

                if (afterChild.Exists) {
                    afterElement.Append(afterChild);
                }

                sawBreak = true;
                continue;
            }

            if (sawBreak) {
                afterElement.Append(child.Clone());
            } else {
                beforeElement.Append(child.Clone());
            }
        }

        if (!containsBreak) {
            return false;
        }

        if (HasNativeRenderableOpenXmlContent(beforeElement)) {
            before = beforeElement;
        }

        if (HasNativeRenderableOpenXmlContent(afterElement)) {
            after = afterElement;
        }

        return true;
    }

    private static bool HasNativeRenderableOpenXmlContent(OfficeXmlNode element) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(WordXml.W, "t")) {
                if (!string.IsNullOrEmpty(node.InnerText)) return true;
            } else if (node.Is(WordXml.W, "tab")) {
                // A tab character, not a tab stop of the paragraph properties.
                if (!node.Parent.IsW("tabs")) return true;
            } else if (node.Is(WordXml.W, "bookmarkStart") || node.Is(WordXml.W, "drawing") ||
                       node.Is(WordXml.W, "footnoteReference") || node.Is(WordXml.W, "endnoteReference") ||
                       node.Is(WordXml.W, "br") || node.Is(WordXml.M, "oMath") || node.Is(WordXml.V, "shape")) {
                return true;
            }
        }

        return false;
    }

    // ---- SDK shims: the members below bridge to files still keyed by SDK elements, and go when those are ported. ----





}
