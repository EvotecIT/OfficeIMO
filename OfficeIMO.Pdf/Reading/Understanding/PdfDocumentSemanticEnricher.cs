using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Document-wide evidence fusion for the canonical semantic page analyses.</summary>
internal static partial class PdfDocumentSemanticEnricher {
    internal static IReadOnlyList<PdfUnderstandingPageResult> Enrich(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        int maxElementsPerPage,
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
        IReadOnlyList<PdfUnderstandingTableCandidate>[] tableCandidates = pages
            .Select(static page => page.TableCandidates)
            .ToArray();
        IReadOnlyList<PdfUnderstandingImageRegion>[] imageRegions = pages
            .Select(static page => page.ImageRegions)
            .ToArray();
        ApplyRepeatedPageEdgeEvidence(document, pageNumbers, pages, elements, workBudget);
        EnsureElementLimits(elements, maxElementsPerPage);
        ApplyOutlineEvidence(document.Outlines, pages, elements, workBudget);
        EnsureElementLimits(elements, maxElementsPerPage);
        TaggedStructureGraph? taggedGraph = BuildTaggedStructureGraph(document, workBudget);
        TaggedContentRoleIndex? taggedRoles = BuildTaggedContentRoleIndex(document, taggedGraph, workBudget);
        TaggedFigureCaptionIndex? taggedFigureCaptions = imageRegions.Any(static regions => regions.Count > 0)
            ? BuildTaggedFigureCaptionIndex(document, taggedGraph, workBudget)
            : null;
        ApplyTaggedTableEvidence(
            document,
            pageNumbers,
            pages,
            tableCandidates,
            taggedGraph,
            maxElementsPerPage,
            workBudget);
        ApplyTaggedStructureEvidence(document, pageNumbers, pages, elements, taggedRoles, workBudget);
        ApplyTaggedTableHeaderEvidence(document, pageNumbers, tableCandidates, taggedRoles, workBudget);
        ApplyTaggedImageEvidence(document, pageNumbers, imageRegions, taggedRoles, workBudget);
        ApplyTaggedFigureCaptionEvidence(document, pageNumbers, imageRegions, elements, taggedFigureCaptions, workBudget);
        EnsureElementLimits(elements, maxElementsPerPage);
        ApplyHeadingFontTierEvidence(elements, workBudget);

        var result = new PdfUnderstandingPageResult[pages.Count];
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            workBudget.Consume();
            PdfUnderstandingPageResult page = pages[pageIndex];
            PdfUnderstandingRegion[] canonicalRegions = elements[pageIndex]
                .Select(static element => element.Region)
                .ToArray();
            bool regionsChanged = canonicalRegions.Length != page.ReadingOrder.Count;
            for (int regionIndex = 0; !regionsChanged && regionIndex < canonicalRegions.Length; regionIndex++) {
                workBudget.Consume();
                regionsChanged = !ReferenceEquals(canonicalRegions[regionIndex], page.ReadingOrder[regionIndex]);
            }
            IReadOnlyList<PdfUnderstandingRegion> regions = regionsChanged
                ? Array.AsReadOnly(canonicalRegions)
                : page.Regions;
            IReadOnlyList<PdfUnderstandingRegion> readingOrder = regionsChanged
                ? regions
                : page.ReadingOrder;
            IReadOnlyList<PdfReadingOrderEvidence> readingOrderEvidence = regionsChanged
                ? BuildReadingOrderEvidence(page.ReadingOrderEvidence, canonicalRegions, workBudget)
                : page.ReadingOrderEvidence;
            result[pageIndex] = new PdfUnderstandingPageResult(
                page.PageNumber,
                page.DecodedRuns,
                page.Words,
                page.Lines,
                regions,
                readingOrder,
                readingOrderEvidence,
                elements[pageIndex].AsReadOnly(),
                page.Trace,
                page.ConsumeWork,
                page.CancellationCheck,
                page.CompleteOperation,
                page.LogicalProjectionLines,
                page.RestrictLogicalProjectionToReadingOrder,
                tableCandidates[pageIndex],
                page.ImagePlacements,
                RemapImageCaptions(imageRegions[pageIndex], elements[pageIndex]));
        }
        return Array.AsReadOnly(result);
    }

    private static IReadOnlyList<PdfUnderstandingImageRegion> RemapImageCaptions(
        IReadOnlyList<PdfUnderstandingImageRegion> imageRegions,
        List<PdfUnderstandingSemanticElement> elements) {
        if (imageRegions.Count == 0) return imageRegions;
        var semanticByRegion = new Dictionary<PdfUnderstandingRegion, PdfUnderstandingSemanticElement>();
        for (int index = 0; index < elements.Count; index++) semanticByRegion[elements[index].Region] = elements[index];
        var result = new PdfUnderstandingImageRegion[imageRegions.Count];
        for (int index = 0; index < imageRegions.Count; index++) {
            PdfUnderstandingImageRegion imageRegion = imageRegions[index];
            PdfUnderstandingSemanticElement? caption = imageRegion.Caption;
            bool captionRejected = false;
            if (caption is not null) {
                if (semanticByRegion.TryGetValue(caption.Region, out PdfUnderstandingSemanticElement? enriched) &&
                    enriched.Kind == PdfUnderstandingSemanticKind.Caption) {
                    caption = enriched;
                } else {
                    caption = null;
                    captionRejected = true;
                }
            }
            IReadOnlyList<PdfInferenceEvidence> evidence = captionRejected
                ? imageRegion.Evidence.Where(static item =>
                    !string.Equals(item.Code, "image-region.caption-proximity", StringComparison.Ordinal) &&
                    !string.Equals(item.Code, "image-region.paint-order", StringComparison.Ordinal)).ToArray()
                : imageRegion.Evidence;
            bool taggedFigure = evidence.Any(static item =>
                string.Equals(item.Code, "image-region.tagged-figure", StringComparison.Ordinal));
            double confidence = captionRejected && !taggedFigure
                ? Math.Min(imageRegion.Confidence, imageRegion.Placement.MarkedContentId.HasValue ? 0.9D : 0.75D)
                : imageRegion.Confidence;
            result[index] = new PdfUnderstandingImageRegion(
                imageRegion.Placement,
                caption,
                confidence,
                evidence,
                imageRegion.IsFigure && (!captionRejected || taggedFigure),
                imageRegion.AlternativeText);
        }
        return Array.AsReadOnly(result);
    }

    private static void EnsureElementLimits(
        IReadOnlyList<PdfUnderstandingSemanticElement>[] elements,
        int maximum) {
        for (int pageIndex = 0; pageIndex < elements.Length; pageIndex++) {
            int actual = elements[pageIndex].Count;
            if (actual > maximum) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, actual);
            }
        }
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
                    PdfUnderstandingSemanticKind.Heading or
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
                double headerBandBottom = height * 0.15D;
                double footerBandTop = height * 0.85D;
                PageEdge edge = visual.Top <= headerBandBottom && visual.Bottom <= headerBandBottom
                    ? PageEdge.Header
                    : visual.Top >= footerBandTop && visual.Bottom >= footerBandTop
                        ? PageEdge.Footer
                        : PageEdge.None;
                bool requiresExactSignature = element.Kind == PdfUnderstandingSemanticKind.Heading;
                string signature = requiresExactSignature
                    ? PdfTextSimilarity.NormalizeSignaturePreservingDigits(region.Text)
                    : PdfTextSimilarity.NormalizeSignature(region.Text);
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

        int uniquePageCount = pageNumbers.Distinct().Count();
        int minimumPages = Math.Max(2, (int)Math.Ceiling(uniquePageCount * 0.5D));
        bool[] repeated = FindRepeatedPageEdgeCandidates(candidates, minimumPages, workBudget);
        for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
            if (!repeated[candidateIndex]) continue;
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
                    "Normalized geometry and exact digit-insensitive text signatures repeat across document pages.",
                    0.95D));
        }
    }

    private static bool[] FindRepeatedPageEdgeCandidates(
        IReadOnlyList<PageEdgeCandidate> candidates,
        int minimumPages,
        PdfUnderstandingWorkBudget workBudget) {
        var bySignature = new Dictionary<PageEdgeSignature, List<int>>();
        for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
            workBudget.Consume();
            PageEdgeCandidate candidate = candidates[candidateIndex];
            var key = new PageEdgeSignature(candidate.Edge, candidate.Signature);
            if (!bySignature.TryGetValue(key, out List<int>? indexes)) {
                indexes = new List<int>();
                bySignature.Add(key, indexes);
            }
            indexes.Add(candidateIndex);
        }

        var repeated = new bool[candidates.Count];
        var pages = new HashSet<int>();
        var compatible = new List<PageEdgeGeometryGroup>();
        foreach (List<int> signatureIndexes in bySignature.Values) {
            workBudget.ThrowIfCancellationRequested();
            pages.Clear();
            for (int index = 0; index < signatureIndexes.Count; index++) {
                pages.Add(candidates[signatureIndexes[index]].PageNumber);
            }
            if (pages.Count < minimumPages) continue;

            PageEdgeGeometryGroup[] geometryGroups = GroupPageEdgeGeometry(candidates, signatureIndexes, workBudget);
            for (int anchorIndex = 0; anchorIndex < geometryGroups.Length; anchorIndex++) {
                PageEdgeGeometryGroup anchor = geometryGroups[anchorIndex];
                pages.Clear();
                compatible.Clear();
                for (int groupIndex = 0; groupIndex < geometryGroups.Length; groupIndex++) {
                    workBudget.Consume();
                    PageEdgeGeometryGroup candidate = geometryGroups[groupIndex];
                    if (Math.Abs(anchor.NormalizedLeft - candidate.NormalizedLeft) > 0.08D ||
                        Math.Abs(anchor.NormalizedWidth - candidate.NormalizedWidth) > 0.12D) continue;
                    compatible.Add(candidate);
                    pages.UnionWith(candidate.PageNumbers);
                }
                if (pages.Count < minimumPages) continue;
                for (int groupIndex = 0; groupIndex < compatible.Count; groupIndex++) {
                    List<int> indexes = compatible[groupIndex].CandidateIndexes;
                    for (int index = 0; index < indexes.Count; index++) repeated[indexes[index]] = true;
                }
            }
        }
        return repeated;
    }

    private static PageEdgeGeometryGroup[] GroupPageEdgeGeometry(
        IReadOnlyList<PageEdgeCandidate> candidates,
        List<int> indexes,
        PdfUnderstandingWorkBudget workBudget) {
        const double geometryPrecision = 1_000_000D;
        var groups = new Dictionary<PageEdgeGeometryKey, PageEdgeGeometryGroup>();
        for (int index = 0; index < indexes.Count; index++) {
            workBudget.Consume();
            int candidateIndex = indexes[index];
            PageEdgeCandidate candidate = candidates[candidateIndex];
            var key = new PageEdgeGeometryKey(
                checked((long)Math.Round(candidate.NormalizedLeft * geometryPrecision, MidpointRounding.AwayFromZero)),
                checked((long)Math.Round(candidate.NormalizedWidth * geometryPrecision, MidpointRounding.AwayFromZero)));
            if (!groups.TryGetValue(key, out PageEdgeGeometryGroup? group)) {
                group = new PageEdgeGeometryGroup(candidate.NormalizedLeft, candidate.NormalizedWidth);
                groups.Add(key, group);
            }
            group.CandidateIndexes.Add(candidateIndex);
            group.PageNumbers.Add(candidate.PageNumber);
        }
        return groups.Values.ToArray();
    }

    private static void ApplyOutlineEvidence(
        IReadOnlyList<PdfOutlineItem> outlines,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        List<PdfUnderstandingSemanticElement>[] elements,
        PdfUnderstandingWorkBudget workBudget) {
        IReadOnlyDictionary<int, IReadOnlyList<PdfOutlineItem>> outlinesByPage =
            IndexOutlinesByPage(outlines, workBudget);
        if (outlinesByPage.Count == 0) return;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            if (!outlinesByPage.TryGetValue(pages[pageIndex].PageNumber, out IReadOnlyList<PdfOutlineItem>? pageOutlines)) continue;
            var usedLines = new HashSet<(long BaselineY, long XStart, string Text)>();
            var normalizedLineText = new Dictionary<PdfUnderstandingLine, string>();
            var normalizedTextScalars = new Dictionary<string, int[]>(StringComparer.Ordinal);
            foreach (PdfOutlineItem outline in pageOutlines) {
                string outlineText = PdfTextSimilarity.NormalizeSignaturePreservingDigits(outline.Title);
                if (outlineText.Length == 0) continue;
                if (!normalizedTextScalars.TryGetValue(outlineText, out int[]? outlineScalars)) {
                    outlineScalars = PdfTextSimilarity.GetScalars(outlineText);
                    normalizedTextScalars.Add(outlineText, outlineScalars);
                }
                OutlineLineCandidate? candidate = FindOutlineLineCandidate(
                    elements[pageIndex],
                    outline,
                    outlineText,
                    outlineScalars,
                    usedLines,
                    normalizedLineText,
                    normalizedTextScalars,
                    requireExactText: true,
                    workBudget);
                candidate ??= FindOutlineLineCandidate(
                    elements[pageIndex],
                    outline,
                    outlineText,
                    outlineScalars,
                    usedLines,
                    normalizedLineText,
                    normalizedTextScalars,
                    requireExactText: false,
                    workBudget);
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
                    List<PdfUnderstandingSemanticElement> split = SplitElementByLine(
                        candidate.Element,
                        candidate.Line,
                        new PdfUnderstandingSemanticElement(
                            CreateLineRegion(candidate.Element.Region, candidate.Line),
                            PdfUnderstandingSemanticKind.Heading,
                            Math.Max(candidate.Element.Confidence, 0.96D),
                            candidate.Element.Evidence.Concat(new[] { evidence }),
                            outline.Level),
                        workBudget);
                    elements[pageIndex].RemoveAt(candidate.ElementIndex);
                    elements[pageIndex].InsertRange(candidate.ElementIndex, split);
                }
            }
        }
    }

    private static OutlineLineCandidate? FindOutlineLineCandidate(
        List<PdfUnderstandingSemanticElement> elements,
        PdfOutlineItem outline,
        string outlineText,
        int[] outlineScalars,
        HashSet<(long BaselineY, long XStart, string Text)> usedLines,
        Dictionary<PdfUnderstandingLine, string> normalizedLineText,
        Dictionary<string, int[]> normalizedTextScalars,
        bool requireExactText,
        PdfUnderstandingWorkBudget workBudget) {
        OutlineLineCandidate? candidate = null;
        for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
            PdfUnderstandingSemanticElement element = elements[elementIndex];
            for (int lineIndex = 0; lineIndex < element.Region.Lines.Count; lineIndex++) {
                workBudget.Consume();
                PdfUnderstandingLine line = element.Region.Lines[lineIndex];
                (long BaselineY, long XStart, string Text) key = CreateLineKey(line);
                if (usedLines.Contains(key)) continue;
                if (!normalizedLineText.TryGetValue(line, out string? lineText)) {
                    lineText = PdfTextSimilarity.NormalizeSignaturePreservingDigits(line.Text);
                    normalizedLineText.Add(line, lineText);
                }
                double score;
                if (requireExactText) {
                    if (!string.Equals(lineText, outlineText, StringComparison.Ordinal)) continue;
                    score = 1D;
                } else {
                    if (!normalizedTextScalars.TryGetValue(lineText, out int[]? lineScalars)) {
                        lineScalars = PdfTextSimilarity.GetScalars(lineText);
                        normalizedTextScalars.Add(lineText, lineScalars);
                    }
                    if (!PdfTextSimilarity.TryGetNormalizedSimilarity(
                        lineScalars,
                        outlineScalars,
                        0.9D,
                        out score,
                        workBudget.Consume,
                        workBudget.ThrowIfCancellationRequested)) {
                        continue;
                    }
                }
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
        return candidate;
    }

    private static (long BaselineY, long XStart, string Text) CreateLineKey(PdfUnderstandingLine line) =>
        (BitConverter.DoubleToInt64Bits(line.BaselineY), BitConverter.DoubleToInt64Bits(line.XStart), line.Text);

    private static TaggedContentRoleIndex? BuildTaggedContentRoleIndex(
        PdfReadDocument document,
        TaggedStructureGraph? graph,
        PdfUnderstandingWorkBudget workBudget) {
        if (graph is null) return null;
        PdfTaggedContentInfo tagged = graph.Tagged;

        var index = new TaggedContentRoleIndex();
        foreach (PdfStructureElementInfo structureElement in tagged.StructureElements) {
            if (!graph.ReachableObjectNumbers.Contains(structureElement.ObjectNumber)) continue;
            if (structureElement.MarkedContentReferences.Count == 0) continue;
            TaggedStructureBinding? binding = ResolveTaggedBinding(
                tagged,
                graph.StructuresByObject,
                structureElement,
                workBudget);
            if (!binding.HasValue) continue;
            foreach (PdfMarkedContentReference reference in structureElement.MarkedContentReferences) {
                workBudget.Consume();
                int? pageObjectNumber = reference.PageObjectNumber ?? binding.Value.PageObjectNumber;
                int? pageNumber = pageObjectNumber.HasValue
                    ? document.GetPageNumberForObject(pageObjectNumber.Value)
                    : null;
                if (!pageNumber.HasValue) continue;
                index.Add(
                    pageNumber.Value,
                    new MarkedContentKey(reference.ContentStreamObjectNumber, reference.MarkedContentId),
                    binding.Value.Roles,
                    binding.Value.AlternativeText);
            }
        }
        return index.IsEmpty ? null : index;
    }

    private static void ApplyTaggedImageEvidence(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingImageRegion>[] imageRegions,
        TaggedContentRoleIndex? taggedRoles,
        PdfUnderstandingWorkBudget workBudget) {
        if (taggedRoles is null) return;
        for (int pageIndex = 0; pageIndex < imageRegions.Length; pageIndex++) {
            int pageNumber = pageNumbers[pageIndex];
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            PdfUnderstandingImageRegion[] current = imageRegions[pageIndex].ToArray();
            for (int imageIndex = 0; imageIndex < current.Length; imageIndex++) {
                workBudget.Consume();
                PdfUnderstandingImageRegion region = current[imageIndex];
                PdfImagePlacement placement = region.Placement;
                if (!placement.MarkedContentId.HasValue) continue;
                var key = new MarkedContentKey(placement.ContentStreamObjectNumber, placement.MarkedContentId.Value);
                IReadOnlyList<string> roles = taggedRoles.GetRoles(pageNumber, readPage, key);
                if (!roles.Any(static role => string.Equals(role, "Figure", StringComparison.OrdinalIgnoreCase))) continue;
                string? alternativeText = taggedRoles.GetAlternativeText(pageNumber, readPage, key);
                current[imageIndex] = new PdfUnderstandingImageRegion(
                    placement,
                    region.Caption,
                    Math.Max(region.Confidence, 0.96D),
                    region.Evidence.Concat(new[] {
                        new PdfInferenceEvidence(
                            "image-region.tagged-figure",
                            "A tagged-PDF Figure structure element owns the image marked content.",
                            0.98D)
                    }),
                    isFigure: true,
                    alternativeText: alternativeText ?? region.AlternativeText);
            }
            imageRegions[pageIndex] = Array.AsReadOnly(current);
        }
    }

    private static void ApplyTaggedStructureEvidence(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        List<PdfUnderstandingSemanticElement>[] elements,
        TaggedContentRoleIndex? taggedRoles,
        PdfUnderstandingWorkBudget workBudget) {
        if (taggedRoles is null) return;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            int pageNumber = pageNumbers[pageIndex];
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            if (!taggedRoles.HasPage(pageNumber)) continue;
            var enriched = new List<PdfUnderstandingSemanticElement>(elements[pageIndex].Count);
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
                        .SelectMany(GetRoles)
                        .Select(TryMapTaggedRole)
                        .Where(static role => role.HasValue)
                        .OrderBy(static role => role!.Value.Priority)
                        .ThenBy(static role => role!.Value.Level ?? int.MaxValue)
                        .FirstOrDefault();
                    if (!bestRole.HasValue) continue;
                    matches.Add(new TaggedLineMatch(line, markedContent, bestRole.Value));
                }
                if (matches.Count == 0) {
                    enriched.Add(current);
                    continue;
                }

                bool coversWholeRegion = matches.Count == current.Region.Lines.Count &&
                    matches.All(match => match.Role.Kind == matches[0].Role.Kind && match.Role.Level == matches[0].Role.Level);
                if (coversWholeRegion) {
                    enriched.Add(WithTaggedEvidence(current, matches[0]));
                    continue;
                }
                var matchByLine = matches.ToDictionary(static match => match.Line);
                for (int lineIndex = 0; lineIndex < current.Region.Lines.Count; lineIndex++) {
                    workBudget.Consume();
                    PdfUnderstandingLine line = current.Region.Lines[lineIndex];
                    PdfUnderstandingSemanticElement lineElement = CreateSplitLineElement(current, line);
                    enriched.Add(matchByLine.TryGetValue(line, out TaggedLineMatch match)
                        ? WithTaggedEvidence(lineElement, match)
                        : lineElement);
                }
            }
            elements[pageIndex] = enriched;

            IEnumerable<string> GetRoles(MarkedContentKey key) {
                foreach (string role in taggedRoles.GetRoles(pageNumber, readPage, key)) yield return role;
            }
        }
    }

    private static void ApplyTaggedTableHeaderEvidence(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingTableCandidate>[] tableCandidates,
        TaggedContentRoleIndex? taggedRoles,
        PdfUnderstandingWorkBudget workBudget) {
        if (taggedRoles is null) return;
        for (int pageIndex = 0; pageIndex < tableCandidates.Length; pageIndex++) {
            int pageNumber = pageNumbers[pageIndex];
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            IReadOnlyList<PdfUnderstandingTableCandidate> source = tableCandidates[pageIndex];
            PdfUnderstandingTableCandidate[]? updated = null;
            for (int tableIndex = 0; tableIndex < source.Count; tableIndex++) {
                workBudget.Consume();
                PdfUnderstandingTableCandidate candidate = source[tableIndex];
                if (candidate.SourceKind != PdfLogicalContentSourceKind.Native ||
                    candidate.Rows.Count < 2 ||
                    !HasTaggedHeaderRow(candidate, pageNumber, readPage, taggedRoles, workBudget)) continue;
                updated ??= source.ToArray();
                updated[tableIndex] = candidate.WithAdditionalEvidence(
                    new PdfInferenceEvidence(
                        "table.tagged-header-row",
                        "Tagged-PDF THead or top-row TH structure identifies the first table row as column headers.",
                        0.99D),
                    0.98D);
            }
            if (updated is not null) tableCandidates[pageIndex] = Array.AsReadOnly(updated);
        }
    }

    private static bool HasTaggedHeaderRow(
        PdfUnderstandingTableCandidate candidate,
        int pageNumber,
        PdfReadPage readPage,
        TaggedContentRoleIndex taggedRoles,
        PdfUnderstandingWorkBudget workBudget) {
        var topRowKeys = new HashSet<MarkedContentKey>();
        double topBaseline = candidate.SourceLines.Count == 0
            ? double.NaN
            : candidate.SourceLines.Max(static line => line.BaselineY);
        double topTolerance = candidate.SourceLines.Count == 0
            ? 0D
            : Math.Max(1D, candidate.SourceLines.Max(static line => line.FontSize) * 0.35D);
        for (int lineIndex = 0; lineIndex < candidate.SourceLines.Count; lineIndex++) {
            PdfUnderstandingLine line = candidate.SourceLines[lineIndex];
            bool belongsToTopRow = Math.Abs(line.BaselineY - topBaseline) <= topTolerance;
            for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                IReadOnlyList<PdfTextSpan> runs = line.Words[wordIndex].SourceRuns;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    workBudget.Consume();
                    PdfTextSpan run = runs[runIndex];
                    if (!run.MarkedContentId.HasValue) continue;
                    var key = new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value);
                    IReadOnlyList<string> roles = taggedRoles.GetRoles(pageNumber, readPage, key);
                    if (roles.Any(static role => string.Equals(role, "THead", StringComparison.OrdinalIgnoreCase))) return true;
                    if (belongsToTopRow &&
                        roles.Any(static role => string.Equals(role, "TH", StringComparison.OrdinalIgnoreCase))) {
                        topRowKeys.Add(key);
                    }
                }
            }
        }
        return topRowKeys.Count >= candidate.Columns.Count;
    }

    private static List<PdfUnderstandingSemanticElement> SplitElementByLine(
        PdfUnderstandingSemanticElement source,
        PdfUnderstandingLine replacementLine,
        PdfUnderstandingSemanticElement replacement,
        PdfUnderstandingWorkBudget workBudget) {
        var split = new List<PdfUnderstandingSemanticElement>(source.Region.Lines.Count);
        for (int lineIndex = 0; lineIndex < source.Region.Lines.Count; lineIndex++) {
            workBudget.Consume();
            PdfUnderstandingLine line = source.Region.Lines[lineIndex];
            split.Add(ReferenceEquals(line, replacementLine)
                ? replacement
                : CreateSplitLineElement(source, line));
        }
        return split;
    }

    private static PdfUnderstandingSemanticElement CreateSplitLineElement(
        PdfUnderstandingSemanticElement source,
        PdfUnderstandingLine line) {
        if (!ContentStructureExtractor.IsListItemText(line.Text)) {
            return new PdfUnderstandingSemanticElement(
                CreateLineRegion(source.Region, line),
                source.Kind,
                source.Confidence,
                source.Evidence,
                source.Level);
        }

        IEnumerable<PdfInferenceEvidence> evidence = source.Evidence;
        if (!source.Evidence.Any(static item =>
                string.Equals(item.Code, "semantic.list-marker", StringComparison.Ordinal))) {
            evidence = evidence.Concat(new[] {
                new PdfInferenceEvidence(
                    "semantic.list-marker",
                    "The split line begins with a bullet or numbered marker.",
                    0.4D)
            });
        }
        return new PdfUnderstandingSemanticElement(
            CreateLineRegion(source.Region, line),
            PdfUnderstandingSemanticKind.ListItem,
            Math.Max(source.Confidence, 0.9D),
            evidence);
    }

    private static PdfUnderstandingRegion CreateLineRegion(
        PdfUnderstandingRegion source,
        PdfUnderstandingLine line) =>
        new(new[] { line }, line.Confidence, source.Evidence);

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfReadingOrderEvidence> BuildReadingOrderEvidence(
        IReadOnlyList<PdfReadingOrderEvidence> source,
        PdfUnderstandingRegion[] regions,
        PdfUnderstandingWorkBudget workBudget) {
        var evidenceByLine = new Dictionary<PdfUnderstandingLine, PdfReadingOrderEvidence>();
        for (int evidenceIndex = 0; evidenceIndex < source.Count; evidenceIndex++) {
            PdfReadingOrderEvidence evidence = source[evidenceIndex];
            for (int lineIndex = 0; lineIndex < evidence.Region.Lines.Count; lineIndex++) {
                workBudget.Consume();
                evidenceByLine[evidence.Region.Lines[lineIndex]] = evidence;
            }
        }
        var result = new PdfReadingOrderEvidence[regions.Length];
        for (int regionIndex = 0; regionIndex < regions.Length; regionIndex++) {
            workBudget.Consume();
            PdfUnderstandingRegion region = regions[regionIndex];
            PdfReadingOrderEvidence? sourceEvidence = region.Lines.Count > 0 &&
                evidenceByLine.TryGetValue(region.Lines[0], out PdfReadingOrderEvidence? matched)
                    ? matched
                    : null;
            result[regionIndex] = new PdfReadingOrderEvidence(
                regionIndex,
                region,
                sourceEvidence?.Confidence ?? region.Confidence,
                sourceEvidence?.Evidence ?? Array.Empty<PdfInferenceEvidence>());
        }
        return Array.AsReadOnly(result);
    }

    private static void ApplyHeadingFontTierEvidence(List<PdfUnderstandingSemanticElement>[] elements, PdfUnderstandingWorkBudget workBudget) {
        var candidates = new List<(int PageIndex, int ElementIndex, double FontSize)>();
        var fontSizes = new List<double>();
        for (int pageIndex = 0; pageIndex < elements.Length; pageIndex++) {
            for (int elementIndex = 0; elementIndex < elements[pageIndex].Count; elementIndex++) {
                workBudget.Consume();
                PdfUnderstandingSemanticElement element = elements[pageIndex][elementIndex];
                if (element.Kind != PdfUnderstandingSemanticKind.Heading ||
                    element.Level.HasValue ||
                    element.Region.Lines.Count == 0) continue;
                double fontSize = element.Region.Lines.Max(static line => line.FontSize);
                candidates.Add((
                    pageIndex,
                    elementIndex,
                    fontSize));
                fontSizes.Add(fontSize);
            }
        }
        if (candidates.Count == 0) return;

        Dictionary<double, int> tierByFontSize = PdfHeadingFontTierAnalysis.BuildLookup(fontSizes, () => workBudget.Consume());

        foreach ((int pageIndex, int elementIndex, double fontSize) in candidates) {
            workBudget.Consume();
            int level = Math.Min(6, tierByFontSize[fontSize]);
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
        Dictionary<int, PdfStructureElementInfo> structuresByObject,
        PdfStructureElementInfo structureElement,
        PdfUnderstandingWorkBudget workBudget) {
        var visited = new HashSet<int>();
        var roles = new List<string>();
        PdfStructureElementInfo? current = structureElement;
        int? pageObjectNumber = null;
        string? alternativeText = null;
        while (current is not null && visited.Add(current.ObjectNumber)) {
            workBudget.Consume();
            pageObjectNumber ??= current.PageObjectNumber;
            if (!string.IsNullOrWhiteSpace(current.StructureType)) {
                string candidate = ResolveRole(tagged, current.StructureType!);
                if (!roles.Contains(candidate, StringComparer.OrdinalIgnoreCase)) roles.Add(candidate);
                if (alternativeText is null &&
                    string.Equals(candidate, "Figure", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(current.AlternateText)) {
                    alternativeText = current.AlternateText;
                }
            }
            current = current.ParentObjectNumber.HasValue &&
                structuresByObject.TryGetValue(current.ParentObjectNumber.Value, out PdfStructureElementInfo? parent)
                ? parent
                : null;
        }
        return roles.Count == 0
            ? null
            : new TaggedStructureBinding(roles.AsReadOnly(), pageObjectNumber, alternativeText);
    }

    private static TaggedRole? TryMapTaggedRole(string role) {
        int? headingLevel = HeadingLevel(role);
        if (headingLevel.HasValue) return new TaggedRole(role, PdfUnderstandingSemanticKind.Heading, headingLevel, 0);
        if (string.Equals(role, "H", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Heading, null, 0);
        if (string.Equals(role, "P", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Paragraph, null, 1);
        if (string.Equals(role, "LI", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.ListItem, null, 2);
        if (string.Equals(role, "Table", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Table, null, 3);
        if (string.Equals(role, "Caption", StringComparison.OrdinalIgnoreCase)) return new TaggedRole(role, PdfUnderstandingSemanticKind.Caption, null, 0);
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

    internal static int? HeadingLevel(string role) =>
        role.Length == 2 && (role[0] == 'H' || role[0] == 'h') && char.IsDigit(role[1])
            && role[1] >= '1' && role[1] <= '6'
            ? role[1] - '0'
            : null;

    private static IEnumerable<PdfOutlineItem> FlattenOutlines(IReadOnlyList<PdfOutlineItem> outlines) {
        var stack = new Stack<(IReadOnlyList<PdfOutlineItem> Items, int Index)>();
        stack.Push((outlines, 0));
        while (stack.Count > 0) {
            (IReadOnlyList<PdfOutlineItem> items, int index) = stack.Pop();
            if (index >= items.Count) continue;
            PdfOutlineItem current = items[index];
            stack.Push((items, index + 1));
            yield return current;
            if (current.Children.Count > 0) stack.Push((current.Children, 0));
        }
    }

    internal static IReadOnlyDictionary<int, IReadOnlyList<PdfOutlineItem>> IndexOutlinesByPage(
        IReadOnlyList<PdfOutlineItem> outlines,
        PdfUnderstandingWorkBudget workBudget) {
        var mutable = new Dictionary<int, List<PdfOutlineItem>>();
        foreach (PdfOutlineItem outline in FlattenOutlines(outlines)) {
            workBudget.Consume();
            if (!outline.PageNumber.HasValue) continue;
            if (!mutable.TryGetValue(outline.PageNumber.Value, out List<PdfOutlineItem>? pageOutlines)) {
                pageOutlines = new List<PdfOutlineItem>();
                mutable.Add(outline.PageNumber.Value, pageOutlines);
            }
            pageOutlines.Add(outline);
        }

        return mutable.ToDictionary(
            static pair => pair.Key,
            static pair => (IReadOnlyList<PdfOutlineItem>)pair.Value.AsReadOnly());
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

    private readonly record struct PageEdgeSignature(PageEdge Edge, string Signature);

    private readonly record struct PageEdgeGeometryKey(long NormalizedLeft, long NormalizedWidth);

    private sealed class PageEdgeGeometryGroup {
        internal PageEdgeGeometryGroup(double normalizedLeft, double normalizedWidth) {
            NormalizedLeft = normalizedLeft;
            NormalizedWidth = normalizedWidth;
        }
        internal double NormalizedLeft { get; }
        internal double NormalizedWidth { get; }
        internal List<int> CandidateIndexes { get; } = new List<int>();
        internal HashSet<int> PageNumbers { get; } = new HashSet<int>();
    }

    private readonly record struct TaggedRole(
        string Name,
        PdfUnderstandingSemanticKind Kind,
        int? Level,
        int Priority);

    private readonly record struct TaggedStructureBinding(
        IReadOnlyList<string> Roles,
        int? PageObjectNumber,
        string? AlternativeText);

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

    private sealed class TaggedContentRoleIndex {
        private readonly Dictionary<int, Dictionary<MarkedContentKey, List<string>>> _rolesByPage = new();
        private readonly Dictionary<int, Dictionary<MarkedContentKey, string>> _alternativeTextByPage = new();
        private readonly Dictionary<int, Dictionary<int, List<MarkedContentKey>>> _scopedKeysByPageAndMcid = new();

        internal bool IsEmpty => _rolesByPage.Count == 0;

        internal void Add(
            int pageNumber,
            MarkedContentKey key,
            IReadOnlyList<string> roles,
            string? alternativeText) {
            if (!_rolesByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, List<string>>? pageRoles)) {
                pageRoles = new Dictionary<MarkedContentKey, List<string>>();
                _rolesByPage.Add(pageNumber, pageRoles);
            }
            if (!pageRoles.TryGetValue(key, out List<string>? values)) {
                values = new List<string>();
                pageRoles.Add(key, values);
            }
            for (int index = 0; index < roles.Count; index++) {
                string role = roles[index];
                if (!values.Contains(role, StringComparer.OrdinalIgnoreCase)) values.Add(role);
            }
            if (key.ContentStreamObjectNumber.HasValue) {
                if (!_scopedKeysByPageAndMcid.TryGetValue(pageNumber, out Dictionary<int, List<MarkedContentKey>>? pageScopedKeys)) {
                    pageScopedKeys = new Dictionary<int, List<MarkedContentKey>>();
                    _scopedKeysByPageAndMcid.Add(pageNumber, pageScopedKeys);
                }
                if (!pageScopedKeys.TryGetValue(key.MarkedContentId, out List<MarkedContentKey>? scopedKeys)) {
                    scopedKeys = new List<MarkedContentKey>();
                    pageScopedKeys.Add(key.MarkedContentId, scopedKeys);
                }
                if (!scopedKeys.Contains(key)) scopedKeys.Add(key);
            }
            if (!string.IsNullOrWhiteSpace(alternativeText)) {
                if (!_alternativeTextByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, string>? pageAlternativeText)) {
                    pageAlternativeText = new Dictionary<MarkedContentKey, string>();
                    _alternativeTextByPage.Add(pageNumber, pageAlternativeText);
                }
#pragma warning disable CA1864 // Dictionary.TryAdd is unavailable for the net472 and netstandard2.0 targets.
                if (!pageAlternativeText.ContainsKey(key)) pageAlternativeText.Add(key, alternativeText!);
#pragma warning restore CA1864
            }
        }

        internal bool HasPage(int pageNumber) => _rolesByPage.ContainsKey(pageNumber);

        internal IReadOnlyList<string> GetRoles(
            int pageNumber,
            PdfReadPage readPage,
            MarkedContentKey key) {
            if (!_rolesByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, List<string>>? pageRoles)) {
                return Array.Empty<string>();
            }
            if (pageRoles.TryGetValue(key, out List<string>? exactRoles)) return exactRoles;
            if (!key.ContentStreamObjectNumber.HasValue) {
                MarkedContentKey? resolved = ResolveUniquePageContentKey(pageNumber, readPage, key.MarkedContentId);
                return resolved.HasValue && pageRoles.TryGetValue(resolved.Value, out List<string>? scopedPageContentRoles)
                    ? scopedPageContentRoles
                    : Array.Empty<string>();
            }
            if (!readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber)) return Array.Empty<string>();
            return pageRoles.TryGetValue(new MarkedContentKey(null, key.MarkedContentId), out List<string>? pageContentRoles)
                ? pageContentRoles
                : Array.Empty<string>();
        }

        internal string? GetAlternativeText(
            int pageNumber,
            PdfReadPage readPage,
            MarkedContentKey key) {
            if (!_alternativeTextByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, string>? pageAlternativeText)) {
                return null;
            }
            if (pageAlternativeText.TryGetValue(key, out string? exact)) return exact;
            if (!key.ContentStreamObjectNumber.HasValue) {
                MarkedContentKey? resolved = ResolveUniquePageContentKey(pageNumber, readPage, key.MarkedContentId);
                return resolved.HasValue && pageAlternativeText.TryGetValue(resolved.Value, out string? scopedPageContent)
                    ? scopedPageContent
                    : null;
            }
            if (!readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber)) return null;
            return pageAlternativeText.TryGetValue(new MarkedContentKey(null, key.MarkedContentId), out string? pageContent)
                ? pageContent
                : null;
        }

        private MarkedContentKey? ResolveUniquePageContentKey(
            int pageNumber,
            PdfReadPage readPage,
            int markedContentId) {
            if (!_scopedKeysByPageAndMcid.TryGetValue(pageNumber, out Dictionary<int, List<MarkedContentKey>>? pageScopedKeys) ||
                !pageScopedKeys.TryGetValue(markedContentId, out List<MarkedContentKey>? candidates)) return null;
            MarkedContentKey? result = null;
            for (int index = 0; index < candidates.Count; index++) {
                MarkedContentKey candidate = candidates[index];
                if (!readPage.IsPageContentStreamObjectNumber(candidate.ContentStreamObjectNumber)) continue;
                if (result.HasValue) return null;
                result = candidate;
            }
            return result;
        }
    }

    private readonly record struct TaggedLineMatch(
        PdfUnderstandingLine Line,
        IReadOnlyList<MarkedContentKey> MarkedContent,
        TaggedRole Role);
}
