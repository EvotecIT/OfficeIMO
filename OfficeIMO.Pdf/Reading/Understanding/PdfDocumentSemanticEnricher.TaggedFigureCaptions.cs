namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    private static TaggedStructureGraph? BuildTaggedStructureGraph(
        PdfReadDocument document,
        PdfUnderstandingWorkBudget workBudget) {
        PdfTaggedContentInfo? tagged = document.TaggedContent;
        if (tagged is null ||
            tagged.StructureElements.Count == 0 ||
            tagged.RootElementObjectNumbers.Count == 0) return null;

        var structuresByObject = new Dictionary<int, PdfStructureElementInfo>(tagged.StructureElements.Count);
        foreach (PdfStructureElementInfo structureElement in tagged.StructureElements) {
            workBudget.Consume();
            structuresByObject.Add(structureElement.ObjectNumber, structureElement);
        }

        var reachable = new HashSet<int>();
        var pending = new Stack<PdfStructureElementInfo>();
        for (int rootIndex = 0; rootIndex < tagged.RootElementObjectNumbers.Count; rootIndex++) {
            workBudget.Consume();
            if (!structuresByObject.TryGetValue(tagged.RootElementObjectNumbers[rootIndex], out PdfStructureElementInfo? root) ||
                root.ParentObjectNumber != tagged.StructTreeRootObjectNumber) continue;
            pending.Push(root);
        }
        while (pending.Count > 0) {
            workBudget.Consume();
            PdfStructureElementInfo current = pending.Pop();
            if (!reachable.Add(current.ObjectNumber)) continue;
            for (int childIndex = 0; childIndex < current.ChildElementObjectNumbers.Count; childIndex++) {
                workBudget.Consume();
                if (structuresByObject.TryGetValue(current.ChildElementObjectNumbers[childIndex], out PdfStructureElementInfo? child) &&
                    child.ParentObjectNumber == current.ObjectNumber) {
                    pending.Push(child);
                }
            }
        }
        return reachable.Count == 0
            ? null
            : new TaggedStructureGraph(tagged, structuresByObject, reachable);
    }

    private static TaggedFigureCaptionIndex? BuildTaggedFigureCaptionIndex(
        PdfReadDocument document,
        TaggedStructureGraph? graph,
        PdfUnderstandingWorkBudget workBudget) {
        if (graph is null) return null;
        PdfTaggedContentInfo tagged = graph.Tagged;
        Dictionary<int, PdfStructureElementInfo> structuresByObject = graph.StructuresByObject;

        var index = new TaggedFigureCaptionIndex();
        foreach (PdfStructureElementInfo figure in tagged.StructureElements) {
            workBudget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(figure.ObjectNumber) ||
                !HasResolvedRole(tagged, figure, "Figure")) continue;
            PdfStructureElementInfo? caption = ResolveAssociatedCaption(graph, figure, workBudget);
            if (caption is null) continue;

            if (!TryCollectStructureMarkedContent(
                document,
                graph,
                figure,
                excludeCaptionSubtrees: true,
                workBudget,
                out PageMarkedContent[] figureContent) ||
                !TryCollectStructureMarkedContent(
                document,
                graph,
                caption,
                excludeCaptionSubtrees: false,
                workBudget,
                out PageMarkedContent[] captionContent)) continue;
            if (figureContent.Length == 0 || captionContent.Length == 0) continue;

            Dictionary<int, HashSet<MarkedContentKey>> figureContentByPage = GroupMarkedContentByPage(figureContent, workBudget);
            Dictionary<int, HashSet<MarkedContentKey>> captionContentByPage = GroupMarkedContentByPage(captionContent, workBudget);
            foreach (KeyValuePair<int, HashSet<MarkedContentKey>> pageFigureContent in figureContentByPage) {
                workBudget.Consume();
                MarkedContentKey[] figureKeys = pageFigureContent.Value.ToArray();
                MarkedContentKey[] captionKeys = captionContentByPage.TryGetValue(pageFigureContent.Key, out HashSet<MarkedContentKey>? pageCaptionContent)
                    ? pageCaptionContent.ToArray()
                    : Array.Empty<MarkedContentKey>();
                workBudget.Consume(figureKeys.Length + captionKeys.Length);
                if (figureKeys.Length == 0 || captionKeys.Length == 0) continue;
                index.Add(pageFigureContent.Key, figureKeys, captionKeys);
            }
        }
        return index.IsEmpty ? null : index;
    }

    private static PdfStructureElementInfo? ResolveAssociatedCaption(
        TaggedStructureGraph graph,
        PdfStructureElementInfo figure,
        PdfUnderstandingWorkBudget workBudget) {
        if (!TryGetChildrenWithRole(
            graph,
            figure,
            "Caption",
            workBudget,
            out List<PdfStructureElementInfo> containedCaptions)) return null;
        if (containedCaptions.Count == 1) return containedCaptions[0];
        if (containedCaptions.Count > 1 ||
            !figure.ParentObjectNumber.HasValue ||
            !graph.StructuresByObject.TryGetValue(figure.ParentObjectNumber.Value, out PdfStructureElementInfo? parent) ||
            !graph.ReachableObjectNumbers.Contains(parent.ObjectNumber)) {
            return null;
        }

        if (!TryGetChildrenWithRole(
            graph,
            parent,
            "Figure",
            workBudget,
            out List<PdfStructureElementInfo> siblingFigures)) return null;
        if (siblingFigures.Count != 1 || siblingFigures[0].ObjectNumber != figure.ObjectNumber) return null;

        if (!TryGetChildrenWithRole(
            graph,
            parent,
            "Caption",
            workBudget,
            out List<PdfStructureElementInfo> siblingCaptions)) return null;
        if (siblingCaptions.Count != 1) return null;
        int captionIndex = IndexOfChild(parent.ChildElementObjectNumbers, siblingCaptions[0].ObjectNumber);
        return captionIndex == 0 || captionIndex == parent.ChildElementObjectNumbers.Count - 1
            ? siblingCaptions[0]
            : null;
    }

    private static bool TryGetChildrenWithRole(
        TaggedStructureGraph graph,
        PdfStructureElementInfo parent,
        string role,
        PdfUnderstandingWorkBudget workBudget,
        out List<PdfStructureElementInfo> result) {
        result = new List<PdfStructureElementInfo>();
        for (int index = 0; index < parent.ChildElementObjectNumbers.Count; index++) {
            workBudget.Consume();
            int objectNumber = parent.ChildElementObjectNumbers[index];
            if (!graph.StructuresByObject.TryGetValue(objectNumber, out PdfStructureElementInfo? child) ||
                child.ParentObjectNumber != parent.ObjectNumber ||
                !graph.ReachableObjectNumbers.Contains(objectNumber)) return false;
            if (HasResolvedRole(graph.Tagged, child, role)) {
                result.Add(child);
            }
        }
        return true;
    }

    private static bool HasResolvedRole(
        PdfTaggedContentInfo tagged,
        PdfStructureElementInfo structureElement,
        string expected) =>
        !string.IsNullOrWhiteSpace(structureElement.StructureType) &&
        string.Equals(ResolveRole(tagged, structureElement.StructureType!), expected, StringComparison.OrdinalIgnoreCase);

    private static int IndexOfChild(IReadOnlyList<int> children, int objectNumber) {
        for (int index = 0; index < children.Count; index++) {
            if (children[index] == objectNumber) return index;
        }
        return -1;
    }

    private static bool TryCollectStructureMarkedContent(
        PdfReadDocument document,
        TaggedStructureGraph graph,
        PdfStructureElementInfo root,
        bool excludeCaptionSubtrees,
        PdfUnderstandingWorkBudget workBudget,
        out PageMarkedContent[] content) {
        var result = new HashSet<PageMarkedContent>();
        var visited = new HashSet<int>();
        var pending = new Stack<PdfStructureElementInfo>();
        pending.Push(root);
        while (pending.Count > 0) {
            workBudget.Consume();
            PdfStructureElementInfo current = pending.Pop();
            if (!graph.ReachableObjectNumbers.Contains(current.ObjectNumber) ||
                !visited.Add(current.ObjectNumber)) {
                content = Array.Empty<PageMarkedContent>();
                return false;
            }
            if (excludeCaptionSubtrees &&
                current.ObjectNumber != root.ObjectNumber &&
                HasResolvedRole(graph.Tagged, current, "Caption")) {
                continue;
            }

            TaggedStructureBinding? binding = ResolveTaggedBinding(
                graph.Tagged,
                graph.StructuresByObject,
                current,
                workBudget);
            int? inheritedPageObjectNumber = binding?.PageObjectNumber;
            for (int referenceIndex = 0; referenceIndex < current.MarkedContentReferences.Count; referenceIndex++) {
                workBudget.Consume();
                PdfMarkedContentReference reference = current.MarkedContentReferences[referenceIndex];
                int? pageObjectNumber = reference.PageObjectNumber ?? inheritedPageObjectNumber;
                int? pageNumber = pageObjectNumber.HasValue
                    ? document.GetPageNumberForObject(pageObjectNumber.Value)
                    : null;
                if (!pageNumber.HasValue) continue;
                var item = new PageMarkedContent(
                    pageNumber.Value,
                    new MarkedContentKey(reference.ContentStreamObjectNumber, reference.MarkedContentId));
                result.Add(item);
            }

            for (int childIndex = current.ChildElementObjectNumbers.Count - 1; childIndex >= 0; childIndex--) {
                workBudget.Consume();
                int childObjectNumber = current.ChildElementObjectNumbers[childIndex];
                if (!graph.StructuresByObject.TryGetValue(childObjectNumber, out PdfStructureElementInfo? child) ||
                    child.ParentObjectNumber != current.ObjectNumber ||
                    !graph.ReachableObjectNumbers.Contains(childObjectNumber)) {
                    content = Array.Empty<PageMarkedContent>();
                    return false;
                }
                pending.Push(child);
            }
        }
        content = result.Count == 0 ? Array.Empty<PageMarkedContent>() : result.ToArray();
        return true;
    }

    private static Dictionary<int, HashSet<MarkedContentKey>> GroupMarkedContentByPage(
        PageMarkedContent[] content,
        PdfUnderstandingWorkBudget workBudget) {
        var result = new Dictionary<int, HashSet<MarkedContentKey>>();
        for (int index = 0; index < content.Length; index++) {
            workBudget.Consume();
            PageMarkedContent item = content[index];
            if (!result.TryGetValue(item.PageNumber, out HashSet<MarkedContentKey>? pageContent)) {
                pageContent = new HashSet<MarkedContentKey>();
                result.Add(item.PageNumber, pageContent);
            }
            pageContent.Add(item.Key);
        }
        return result;
    }

    private static void ApplyTaggedFigureCaptionEvidence(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingImageRegion>[] imageRegions,
        List<PdfUnderstandingSemanticElement>[] elements,
        TaggedFigureCaptionIndex? taggedFigureCaptions,
        PdfUnderstandingWorkBudget workBudget) {
        if (taggedFigureCaptions is null) return;
        for (int pageIndex = 0; pageIndex < imageRegions.Length; pageIndex++) {
            workBudget.Consume();
            int pageNumber = pageNumbers[pageIndex];
            IReadOnlyList<TaggedFigureCaptionBinding> bindings = taggedFigureCaptions.GetBindings(pageNumber);
            if (bindings.Count == 0) continue;
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            PdfUnderstandingImageRegion[] updated = imageRegions[pageIndex].ToArray();
            bool changed = false;
            for (int imageIndex = 0; imageIndex < updated.Length; imageIndex++) {
                workBudget.Consume();
                PdfUnderstandingImageRegion region = updated[imageIndex];
                if (!region.Placement.MarkedContentId.HasValue) continue;
                var placementKey = new MarkedContentKey(
                    region.Placement.ContentStreamObjectNumber,
                    region.Placement.MarkedContentId.Value);
                var matchedCaptions = new List<PdfUnderstandingSemanticElement>();
                for (int bindingIndex = 0; bindingIndex < bindings.Count; bindingIndex++) {
                    workBudget.Consume();
                    TaggedFigureCaptionBinding binding = bindings[bindingIndex];
                    if (!ContainsMatchingKey(readPage, binding.FigureContent, placementKey, workBudget)) continue;
                    PdfUnderstandingSemanticElement? caption = FindCaptionElement(
                        readPage,
                        elements[pageIndex],
                        binding.CaptionContent,
                        workBudget);
                    if (caption is not null && !matchedCaptions.Contains(caption)) matchedCaptions.Add(caption);
                }
                if (matchedCaptions.Count != 1) continue;
                updated[imageIndex] = WithTaggedCaption(region, matchedCaptions[0]);
                changed = true;
            }
            if (changed) imageRegions[pageIndex] = Array.AsReadOnly(updated);
        }
    }

    private static PdfUnderstandingSemanticElement? FindCaptionElement(
        PdfReadPage readPage,
        List<PdfUnderstandingSemanticElement> elements,
        IReadOnlyList<MarkedContentKey> captionContent,
        PdfUnderstandingWorkBudget workBudget) {
        PdfUnderstandingSemanticElement? result = null;
        for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
            workBudget.Consume();
            PdfUnderstandingSemanticElement element = elements[elementIndex];
            if (element.Kind != PdfUnderstandingSemanticKind.Caption ||
                !ElementContainsMarkedContent(readPage, element, captionContent, workBudget)) continue;
            if (result is not null) return null;
            result = element;
        }
        return result;
    }

    private static bool ElementContainsMarkedContent(
        PdfReadPage readPage,
        PdfUnderstandingSemanticElement element,
        IReadOnlyList<MarkedContentKey> expected,
        PdfUnderstandingWorkBudget workBudget) {
        for (int lineIndex = 0; lineIndex < element.Region.Lines.Count; lineIndex++) {
            IReadOnlyList<PdfUnderstandingWord> words = element.Region.Lines[lineIndex].Words;
            for (int wordIndex = 0; wordIndex < words.Count; wordIndex++) {
                IReadOnlyList<PdfTextSpan> runs = words[wordIndex].SourceRuns;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    workBudget.Consume();
                    PdfTextSpan run = runs[runIndex];
                    if (!run.MarkedContentId.HasValue) continue;
                    var actual = new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value);
                    if (ContainsMatchingKey(readPage, expected, actual, workBudget)) return true;
                }
            }
        }
        return false;
    }

    private static bool ContainsMatchingKey(
        PdfReadPage readPage,
        IReadOnlyList<MarkedContentKey> expected,
        MarkedContentKey actual,
        PdfUnderstandingWorkBudget workBudget) {
        for (int index = 0; index < expected.Count; index++) {
            workBudget.Consume();
            MarkedContentKey candidate = expected[index];
            if (candidate.MarkedContentId != actual.MarkedContentId) continue;
            if (candidate.ContentStreamObjectNumber == actual.ContentStreamObjectNumber) return true;
            if (!candidate.ContentStreamObjectNumber.HasValue &&
                readPage.IsPageContentStreamObjectNumber(actual.ContentStreamObjectNumber)) return true;
            if (!actual.ContentStreamObjectNumber.HasValue &&
                readPage.IsPageContentStreamObjectNumber(candidate.ContentStreamObjectNumber)) return true;
        }
        return false;
    }

    private static PdfUnderstandingImageRegion WithTaggedCaption(
        PdfUnderstandingImageRegion region,
        PdfUnderstandingSemanticElement caption) {
        IEnumerable<PdfInferenceEvidence> evidence = region.Evidence;
        if (region.Caption is not null && !ReferenceEquals(region.Caption, caption)) {
            evidence = evidence.Where(static item =>
                !string.Equals(item.Code, "image-region.caption-proximity", StringComparison.Ordinal) &&
                !string.Equals(item.Code, "image-region.paint-order", StringComparison.Ordinal));
        }
        if (!evidence.Any(static item => string.Equals(item.Code, "image-region.tagged-caption", StringComparison.Ordinal))) {
            evidence = evidence.Concat(new[] {
                new PdfInferenceEvidence(
                    "image-region.tagged-caption",
                    "An unambiguous tagged-PDF Figure and Caption structure relationship owns the image and caption content.",
                    0.99D)
            });
        }
        return new PdfUnderstandingImageRegion(
            region.Placement,
            caption,
            Math.Max(region.Confidence, 0.99D),
            evidence,
            isFigure: true,
            region.AlternativeText);
    }

    private sealed class TaggedStructureGraph {
        internal TaggedStructureGraph(
            PdfTaggedContentInfo tagged,
            Dictionary<int, PdfStructureElementInfo> structuresByObject,
            HashSet<int> reachableObjectNumbers) {
            Tagged = tagged;
            StructuresByObject = structuresByObject;
            ReachableObjectNumbers = reachableObjectNumbers;
        }

        internal PdfTaggedContentInfo Tagged { get; }
        internal Dictionary<int, PdfStructureElementInfo> StructuresByObject { get; }
        internal HashSet<int> ReachableObjectNumbers { get; }
    }

    private readonly record struct PageMarkedContent(int PageNumber, MarkedContentKey Key);

    private readonly record struct TaggedFigureCaptionBinding(
        IReadOnlyList<MarkedContentKey> FigureContent,
        IReadOnlyList<MarkedContentKey> CaptionContent);

    private sealed class TaggedFigureCaptionIndex {
        private readonly Dictionary<int, List<TaggedFigureCaptionBinding>> _bindingsByPage = new();

        internal bool IsEmpty => _bindingsByPage.Count == 0;

        internal void Add(
            int pageNumber,
            IReadOnlyList<MarkedContentKey> figureContent,
            IReadOnlyList<MarkedContentKey> captionContent) {
            if (!_bindingsByPage.TryGetValue(pageNumber, out List<TaggedFigureCaptionBinding>? bindings)) {
                bindings = new List<TaggedFigureCaptionBinding>();
                _bindingsByPage.Add(pageNumber, bindings);
            }
            bindings.Add(new TaggedFigureCaptionBinding(figureContent, captionContent));
        }

        internal IReadOnlyList<TaggedFigureCaptionBinding> GetBindings(int pageNumber) =>
            _bindingsByPage.TryGetValue(pageNumber, out List<TaggedFigureCaptionBinding>? bindings)
                ? bindings
                : Array.Empty<TaggedFigureCaptionBinding>();
    }
}
