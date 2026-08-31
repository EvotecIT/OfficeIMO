using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Document-wide evidence fusion for the canonical semantic page analyses.</summary>
internal static class PdfDocumentSemanticEnricher {
    internal static IReadOnlyList<PdfUnderstandingPageResult> Enrich(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        long maxWorkUnits,
        CancellationToken cancellationToken) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        Guard.NotNull(pages, nameof(pages));
        if (pageNumbers.Length != pages.Count) {
            throw new ArgumentException("Semantic page count must match the selected page count.", nameof(pages));
        }
        if (pages.Count == 0) return pages;
        var workBudget = new PdfUnderstandingWorkBudget(maxWorkUnits, cancellationToken);

        List<PdfUnderstandingSemanticElement>[] elements = pages
            .Select(static page => page.Elements.ToList())
            .ToArray();
        ApplyRepeatedPageEdgeEvidence(document, pageNumbers, pages, elements, workBudget);
        ApplyOutlineEvidence(document.Outlines, pages, elements, workBudget);
        ApplyTaggedStructureEvidence(document, pageNumbers, pages, elements, workBudget);
        ApplyHeadingFontTierEvidence(elements, workBudget);

        var result = new PdfUnderstandingPageResult[pages.Count];
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            workBudget.Consume();
            PdfUnderstandingPageResult page = pages[pageIndex];
            result[pageIndex] = new PdfUnderstandingPageResult(
                page.PageNumber,
                page.DecodedRuns,
                page.Words,
                page.Lines,
                page.Regions,
                page.ReadingOrder,
                page.ReadingOrderEvidence,
                elements[pageIndex].AsReadOnly(),
                page.Trace);
        }
        return Array.AsReadOnly(result);
    }

    private static void ApplyRepeatedPageEdgeEvidence(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        List<PdfUnderstandingSemanticElement>[] elements,
        PdfUnderstandingWorkBudget workBudget) {
        var candidates = new List<PageEdgeCandidate>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfReadPage readPage = document.Pages[pageNumbers[pageIndex] - 1];
            (double width, double height) = readPage.GetInteractionPageSize();
            if (width <= 0D || height <= 0D) continue;
            for (int elementIndex = 0; elementIndex < elements[pageIndex].Count; elementIndex++) {
                workBudget.Consume();
                PdfUnderstandingSemanticElement element = elements[pageIndex][elementIndex];
                if (element.Kind is not (PdfUnderstandingSemanticKind.Paragraph or
                    PdfUnderstandingSemanticKind.Unknown or
                    PdfUnderstandingSemanticKind.Header or
                    PdfUnderstandingSemanticKind.Footer or
                    PdfUnderstandingSemanticKind.Footnote)) continue;
                PdfUnderstandingRegion region = element.Region;
                PdfVisualBounds visual = readPage.TransformBoundsToVisual(
                    region.XStart,
                    region.YBottom,
                    region.XEnd,
                    region.YTop);
                PageEdge edge = visual.Top <= height * 0.15D
                    ? PageEdge.Header
                    : visual.Bottom >= height * 0.85D
                        ? PageEdge.Footer
                        : PageEdge.None;
                string signature = PdfTextSimilarity.NormalizeSignature(region.Text);
                if (edge == PageEdge.None || signature.Length == 0) continue;
                candidates.Add(new PageEdgeCandidate(
                    pageIndex,
                    elementIndex,
                    pageNumbers[pageIndex],
                    edge,
                    signature,
                    visual.Left / width,
                    Math.Max(0D, visual.Width) / width));
            }
        }
        if (candidates.Count < 2) return;

        int[] parent = Enumerable.Range(0, candidates.Count).ToArray();
        for (int leftIndex = 0; leftIndex < candidates.Count; leftIndex++) {
            PageEdgeCandidate left = candidates[leftIndex];
            for (int rightIndex = leftIndex + 1; rightIndex < candidates.Count; rightIndex++) {
                workBudget.Consume();
                PageEdgeCandidate right = candidates[rightIndex];
                if (left.PageNumber == right.PageNumber || left.Edge != right.Edge) continue;
                if (Math.Abs(left.NormalizedLeft - right.NormalizedLeft) > 0.08D ||
                    Math.Abs(left.NormalizedWidth - right.NormalizedWidth) > 0.12D) continue;
                if (PdfTextSimilarity.NormalizedSimilarity(left.Signature, right.Signature) < 0.8D) continue;
                Union(parent, leftIndex, rightIndex);
            }
        }

        int uniquePageCount = pageNumbers.Distinct().Count();
        int minimumPages = Math.Max(2, (int)Math.Ceiling(uniquePageCount * 0.5D));
        foreach (IGrouping<int, int> cluster in Enumerable.Range(0, candidates.Count).GroupBy(index => Find(parent, index))) {
            int[] indexes = cluster.ToArray();
            if (indexes.Select(index => candidates[index].PageNumber).Distinct().Count() < minimumPages) continue;
            foreach (int candidateIndex in indexes) {
                PageEdgeCandidate candidate = candidates[candidateIndex];
                PdfUnderstandingSemanticElement current = elements[candidate.PageIndex][candidate.ElementIndex];
                PdfUnderstandingSemanticKind kind = candidate.Edge == PageEdge.Header
                    ? PdfUnderstandingSemanticKind.Header
                    : PdfUnderstandingSemanticKind.Footer;
                elements[candidate.PageIndex][candidate.ElementIndex] = WithEvidence(
                    current,
                    kind,
                    Math.Max(current.Confidence, 0.94D),
                    new PdfInferenceEvidence(
                        candidate.Edge == PageEdge.Header ? "semantic.repeated-header" : "semantic.repeated-footer",
                        "Normalized geometry and fuzzy text signatures repeat across document pages.",
                        0.95D));
            }
        }
    }

    private static void ApplyOutlineEvidence(
        IReadOnlyList<PdfOutlineItem> outlines,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        List<PdfUnderstandingSemanticElement>[] elements,
        PdfUnderstandingWorkBudget workBudget) {
        PdfOutlineItem[] flattened = FlattenOutlines(outlines).Where(static outline => outline.PageNumber.HasValue).ToArray();
        if (flattened.Length == 0) return;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfOutlineItem[] pageOutlines = flattened.Where(outline => outline.PageNumber == pages[pageIndex].PageNumber).ToArray();
            if (pageOutlines.Length == 0) continue;
            var usedLines = new HashSet<(long BaselineY, long XStart, string Text)>();
            var lineElements = new List<PdfUnderstandingSemanticElement>();
            foreach (PdfOutlineItem outline in pageOutlines) {
                string outlineText = PdfTextSimilarity.NormalizeSignaturePreservingDigits(outline.Title);
                if (outlineText.Length == 0) continue;
                OutlineLineCandidate? candidate = null;
                for (int elementIndex = 0; elementIndex < elements[pageIndex].Count; elementIndex++) {
                    PdfUnderstandingSemanticElement element = elements[pageIndex][elementIndex];
                    for (int lineIndex = 0; lineIndex < element.Region.Lines.Count; lineIndex++) {
                        workBudget.Consume();
                        PdfUnderstandingLine line = element.Region.Lines[lineIndex];
                        (long BaselineY, long XStart, string Text) key = CreateLineKey(line);
                        if (usedLines.Contains(key)) continue;
                        double score = PdfTextSimilarity.NormalizedSimilarity(
                            PdfTextSimilarity.NormalizeSignaturePreservingDigits(line.Text),
                            outlineText);
                        if (score < 0.9D) continue;
                        var current = new OutlineLineCandidate(
                            element,
                            elementIndex,
                            line,
                            key,
                            score,
                            outline.DestinationTop.HasValue
                                ? Math.Abs(line.BaselineY - outline.DestinationTop.Value)
                                : 0D);
                        if (candidate is null || IsBetterOutlineCandidate(current, candidate)) candidate = current;
                    }
                }
                if (candidate is null) continue;

                usedLines.Add(candidate.Key);
                var evidence = new PdfInferenceEvidence(
                    "semantic.outline-heading",
                    "A page-targeted outline title matches this line at outline level " + outline.Level + ".",
                    0.98D);
                if (candidate.Element.Region.Lines.Count == 1) {
                    elements[pageIndex][candidate.ElementIndex] = WithEvidence(
                        candidate.Element,
                        PdfUnderstandingSemanticKind.Heading,
                        Math.Max(candidate.Element.Confidence, 0.96D),
                        evidence,
                        outline.Level);
                } else {
                    var lineRegion = new PdfUnderstandingRegion(
                        new[] { candidate.Line },
                        candidate.Line.Confidence,
                        candidate.Element.Region.Evidence);
                    lineElements.Add(new PdfUnderstandingSemanticElement(
                        lineRegion,
                        PdfUnderstandingSemanticKind.Heading,
                        Math.Max(candidate.Element.Confidence, 0.96D),
                        candidate.Element.Evidence.Concat(new[] { evidence }),
                        outline.Level));
                }
            }
            if (lineElements.Count > 0) elements[pageIndex].InsertRange(0, lineElements);
        }
    }

    private static (long BaselineY, long XStart, string Text) CreateLineKey(PdfUnderstandingLine line) =>
        (BitConverter.DoubleToInt64Bits(line.BaselineY), BitConverter.DoubleToInt64Bits(line.XStart), line.Text);

    private static void ApplyTaggedStructureEvidence(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        List<PdfUnderstandingSemanticElement>[] elements,
        PdfUnderstandingWorkBudget workBudget) {
        PdfTaggedContentInfo? tagged = document.TaggedContent;
        if (tagged is null || tagged.StructureElements.Count == 0) return;
        IReadOnlyDictionary<int, PdfStructureElementInfo> structuresByObject = tagged.StructureElements
            .ToDictionary(static element => element.ObjectNumber);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            int pageNumber = pageNumbers[pageIndex];
            var rolesByMarkedContent = new Dictionary<MarkedContentKey, List<string>>();
            foreach (PdfStructureElementInfo structureElement in tagged.StructureElements) {
                if (structureElement.MarkedContentReferences.Count == 0) continue;
                TaggedStructureBinding? binding = ResolveTaggedBinding(
                    tagged,
                    structuresByObject,
                    structureElement,
                    workBudget);
                if (!binding.HasValue) continue;
                foreach (PdfMarkedContentReference reference in structureElement.MarkedContentReferences) {
                    workBudget.Consume();
                    int? pageObjectNumber = reference.PageObjectNumber ?? binding.Value.PageObjectNumber;
                    if (!pageObjectNumber.HasValue || document.GetPageNumberForObject(pageObjectNumber.Value) != pageNumber) continue;
                    var key = new MarkedContentKey(reference.ContentStreamObjectNumber, reference.MarkedContentId);
                    if (!rolesByMarkedContent.TryGetValue(key, out List<string>? roles)) {
                        roles = new List<string>();
                        rolesByMarkedContent.Add(key, roles);
                    }
                    if (!roles.Contains(binding.Value.Role, StringComparer.OrdinalIgnoreCase)) roles.Add(binding.Value.Role);
                }
            }
            if (rolesByMarkedContent.Count == 0) continue;
            var lineElements = new List<PdfUnderstandingSemanticElement>();
            for (int elementIndex = 0; elementIndex < elements[pageIndex].Count; elementIndex++) {
                workBudget.Consume();
                PdfUnderstandingSemanticElement current = elements[pageIndex][elementIndex];
                var matches = new List<TaggedLineMatch>();
                for (int lineIndex = 0; lineIndex < current.Region.Lines.Count; lineIndex++) {
                    PdfUnderstandingLine line = current.Region.Lines[lineIndex];
                    MarkedContentKey[] markedContent = line.Words
                        .SelectMany(static word => word.SourceRuns)
                        .Where(static run => run.MarkedContentId.HasValue)
                        .Select(static run => new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId!.Value))
                        .Distinct()
                        .ToArray();
                    if (markedContent.Length == 0) continue;
                    workBudget.Consume(markedContent.Length);
                    TaggedRole? bestRole = markedContent
                        .Where(rolesByMarkedContent.ContainsKey)
                        .SelectMany(key => rolesByMarkedContent[key])
                        .Select(TryMapTaggedRole)
                        .Where(static role => role.HasValue)
                        .OrderBy(static role => role!.Value.Priority)
                        .ThenBy(static role => role!.Value.Level ?? int.MaxValue)
                        .FirstOrDefault();
                    if (!bestRole.HasValue) continue;
                    matches.Add(new TaggedLineMatch(line, markedContent, bestRole.Value));
                }
                if (matches.Count == 0) continue;

                bool coversWholeRegion = matches.Count == current.Region.Lines.Count &&
                    matches.All(match => match.Role.Kind == matches[0].Role.Kind && match.Role.Level == matches[0].Role.Level);
                if (coversWholeRegion) {
                    elements[pageIndex][elementIndex] = WithTaggedEvidence(current, matches[0]);
                    continue;
                }
                for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++) {
                    TaggedLineMatch match = matches[matchIndex];
                    var lineRegion = new PdfUnderstandingRegion(
                        new[] { match.Line },
                        match.Line.Confidence,
                        current.Region.Evidence);
                    var lineElement = new PdfUnderstandingSemanticElement(
                        lineRegion,
                        current.Kind,
                        current.Confidence,
                        current.Evidence,
                        current.Level);
                    lineElements.Add(WithTaggedEvidence(lineElement, match));
                }
            }
            if (lineElements.Count > 0) elements[pageIndex].InsertRange(0, lineElements);
        }
    }

    private static void ApplyHeadingFontTierEvidence(List<PdfUnderstandingSemanticElement>[] elements, PdfUnderstandingWorkBudget workBudget) {
        var candidates = new List<(int PageIndex, int ElementIndex, double FontSize)>();
        for (int pageIndex = 0; pageIndex < elements.Length; pageIndex++) {
            for (int elementIndex = 0; elementIndex < elements[pageIndex].Count; elementIndex++) {
                workBudget.Consume();
                PdfUnderstandingSemanticElement element = elements[pageIndex][elementIndex];
                if (element.Kind != PdfUnderstandingSemanticKind.Heading ||
                    element.Level.HasValue ||
                    element.Region.Lines.Count == 0) continue;
                candidates.Add((
                    pageIndex,
                    elementIndex,
                    element.Region.Lines.Max(static line => line.FontSize)));
            }
        }
        if (candidates.Count == 0) return;

        var tiers = new List<double>();
        foreach (double fontSize in candidates.Select(static candidate => candidate.FontSize).OrderByDescending(static size => size)) {
            if (tiers.Count == 0 || Math.Abs(tiers[tiers.Count - 1] - fontSize) > 0.5D) {
                tiers.Add(fontSize);
            }
        }

        foreach ((int pageIndex, int elementIndex, double fontSize) in candidates) {
            int level = Math.Min(6, tiers.FindIndex(tier => Math.Abs(tier - fontSize) <= 0.5D) + 1);
            PdfUnderstandingSemanticElement current = elements[pageIndex][elementIndex];
            elements[pageIndex][elementIndex] = WithEvidence(
                current,
                current.Kind,
                current.Confidence,
                new PdfInferenceEvidence(
                    "semantic.document-heading-font-tier",
                    "The heading level was ranked from document-wide heading font tiers.",
                    0.8D),
                level);
        }
    }

    private static string ResolveRole(PdfTaggedContentInfo tagged, string role) =>
        tagged.RoleMap.TryGetValue(role, out string? mapped) ? mapped : role;

    private static TaggedStructureBinding? ResolveTaggedBinding(
        PdfTaggedContentInfo tagged,
        IReadOnlyDictionary<int, PdfStructureElementInfo> structuresByObject,
        PdfStructureElementInfo structureElement,
        PdfUnderstandingWorkBudget workBudget) {
        var visited = new HashSet<int>();
        PdfStructureElementInfo? current = structureElement;
        string? semanticRole = null;
        int? pageObjectNumber = null;
        while (current is not null && visited.Add(current.ObjectNumber)) {
            workBudget.Consume();
            pageObjectNumber ??= current.PageObjectNumber;
            if (semanticRole is null && !string.IsNullOrWhiteSpace(current.StructureType)) {
                string candidate = ResolveRole(tagged, current.StructureType!);
                if (TryMapTaggedRole(candidate).HasValue) semanticRole = candidate;
            }
            if (semanticRole is not null && pageObjectNumber.HasValue) break;
            current = current.ParentObjectNumber.HasValue &&
                structuresByObject.TryGetValue(current.ParentObjectNumber.Value, out PdfStructureElementInfo? parent)
                ? parent
                : null;
        }
        return semanticRole is null
            ? null
            : new TaggedStructureBinding(semanticRole, pageObjectNumber);
    }

    private static TaggedRole? TryMapTaggedRole(string role) {
        int? headingLevel = HeadingLevel(role);
        if (headingLevel.HasValue) return new TaggedRole(role, PdfUnderstandingSemanticKind.Heading, headingLevel, 0);
        if (string.Equals(role, "H", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Heading, null, 0);
        if (string.Equals(role, "P", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Paragraph, null, 1);
        if (string.Equals(role, "LI", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.ListItem, null, 2);
        if (string.Equals(role, "Table", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Table, null, 3);
        if (string.Equals(role, "Caption", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Caption, null, 4);
        if (string.Equals(role, "Header", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Header, null, 5);
        if (string.Equals(role, "Footer", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Footer, null, 5);
        return null;
    }

    private static PdfUnderstandingSemanticElement WithTaggedEvidence(
        PdfUnderstandingSemanticElement current,
        TaggedLineMatch match) => WithEvidence(
            current,
            match.Role.Kind,
            Math.Max(current.Confidence, 0.96D),
            new PdfInferenceEvidence(
                "semantic.tagged-pdf-role",
                "Tagged-PDF role " + match.Role.Name + " owns marked content " +
                    string.Join(", ", match.MarkedContent.Select(static item => item.Format())) + ".",
                0.98D),
            match.Role.Level ?? current.Level);

    private static PdfUnderstandingSemanticElement WithEvidence(
        PdfUnderstandingSemanticElement current,
        PdfUnderstandingSemanticKind kind,
        double confidence,
        PdfInferenceEvidence evidence,
        int? level = null) => new PdfUnderstandingSemanticElement(
            current.Region,
            kind,
            confidence,
            current.Evidence.Concat(new[] { evidence }),
            level ?? current.Level);

    private static int? HeadingLevel(string role) =>
        role.Length == 2 && (role[0] == 'H' || role[0] == 'h') && char.IsDigit(role[1])
            && role[1] >= '1' && role[1] <= '6'
            ? role[1] - '0'
            : null;

    private static IEnumerable<PdfOutlineItem> FlattenOutlines(IReadOnlyList<PdfOutlineItem> outlines) {
        var stack = new Stack<PdfOutlineItem>(outlines.Reverse());
        while (stack.Count > 0) {
            PdfOutlineItem current = stack.Pop();
            yield return current;
            for (int childIndex = current.Children.Count - 1; childIndex >= 0; childIndex--) {
                stack.Push(current.Children[childIndex]);
            }
        }
    }

    private static bool IsBetterOutlineCandidate(OutlineLineCandidate current, OutlineLineCandidate previous) {
        int distance = current.DestinationDistance.CompareTo(previous.DestinationDistance);
        if (distance != 0) return distance < 0;
        int fontSize = current.Line.FontSize.CompareTo(previous.Line.FontSize);
        if (fontSize != 0) return fontSize > 0;
        int score = current.Score.CompareTo(previous.Score);
        if (score != 0) return score > 0;
        return current.Element.Region.Lines.Count < previous.Element.Region.Lines.Count;
    }

    private static int Find(int[] parent, int value) {
        while (parent[value] != value) {
            parent[value] = parent[parent[value]];
            value = parent[value];
        }
        return value;
    }

    private static void Union(int[] parent, int left, int right) {
        int leftRoot = Find(parent, left);
        int rightRoot = Find(parent, right);
        if (leftRoot != rightRoot) parent[rightRoot] = leftRoot;
    }

    private enum PageEdge { None, Header, Footer }

    private readonly struct PageEdgeCandidate {
        internal PageEdgeCandidate(int pageIndex, int elementIndex, int pageNumber, PageEdge edge, string signature, double normalizedLeft, double normalizedWidth) {
            PageIndex = pageIndex;
            ElementIndex = elementIndex;
            PageNumber = pageNumber;
            Edge = edge;
            Signature = signature;
            NormalizedLeft = normalizedLeft;
            NormalizedWidth = normalizedWidth;
        }
        internal int PageIndex { get; }
        internal int ElementIndex { get; }
        internal int PageNumber { get; }
        internal PageEdge Edge { get; }
        internal string Signature { get; }
        internal double NormalizedLeft { get; }
        internal double NormalizedWidth { get; }
    }

    private readonly record struct TaggedRole(
        string Name,
        PdfUnderstandingSemanticKind Kind,
        int? Level,
        int Priority);

    private readonly record struct TaggedStructureBinding(string Role, int? PageObjectNumber);

    private sealed class OutlineLineCandidate {
        internal OutlineLineCandidate(
            PdfUnderstandingSemanticElement element,
            int elementIndex,
            PdfUnderstandingLine line,
            (long BaselineY, long XStart, string Text) key,
            double score,
            double destinationDistance) {
            Element = element;
            ElementIndex = elementIndex;
            Line = line;
            Key = key;
            Score = score;
            DestinationDistance = destinationDistance;
        }
        internal PdfUnderstandingSemanticElement Element { get; }
        internal int ElementIndex { get; }
        internal PdfUnderstandingLine Line { get; }
        internal (long BaselineY, long XStart, string Text) Key { get; }
        internal double Score { get; }
        internal double DestinationDistance { get; }
    }

    private readonly record struct MarkedContentKey(int? ContentStreamObjectNumber, int MarkedContentId) {
        internal string Format() => ContentStreamObjectNumber.HasValue
            ? MarkedContentId.ToString(System.Globalization.CultureInfo.InvariantCulture) + " in stream " +
                ContentStreamObjectNumber.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : MarkedContentId.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private readonly record struct TaggedLineMatch(
        PdfUnderstandingLine Line,
        IReadOnlyList<MarkedContentKey> MarkedContent,
        TaggedRole Role);
}
