namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    // A section of direct page-content paragraphs can describe one story across
    // columns. Use it only when its order agrees with paint and it owns all text.
    private static Dictionary<int, Dictionary<int, int>?>? IndexTaggedParagraphStories(
        PdfReadDocument document,
        TaggedStructureGraph? graph,
        PdfUnderstandingWorkBudget budget) {
        if (graph is null) return null;
        Dictionary<int, Dictionary<int, int>?>? stories = null;

        foreach (PdfStructureElementInfo section in graph.Tagged.StructureElements) {
            budget.Consume();
            if (!graph.ReachableObjectNumbers.Contains(section.ObjectNumber) ||
                !HasResolvedRole(graph.Tagged, section, "Sect") ||
                section.ChildElementObjectNumbers.Count < 2 ||
                section.MarkedContentReferenceCount != 0) continue;

            var owners = new Dictionary<int, int>();
            int? pageObject = null;
            for (int index = 0; index < section.ChildElementObjectNumbers.Count; index++) {
                budget.Consume();
                if (!TryGetReciprocalChild(graph, section, section.ChildElementObjectNumbers[index], out var paragraph) ||
                    !HasResolvedRole(graph.Tagged, paragraph!, "P") ||
                    paragraph!.MarkedContentReferenceCount != 1 ||
                    paragraph.HasChildElements) break;
                PdfMarkedContentReference reference = paragraph.MarkedContentReferences[0];
                int? paragraphPage = reference.PageObjectNumber ?? paragraph.PageObjectNumber ?? section.PageObjectNumber;
                if (paragraphPage is null || reference.ContentStreamObjectNumber.HasValue ||
                    (pageObject.HasValue && pageObject != paragraphPage) || owners.ContainsKey(reference.MarkedContentId)) break;
                pageObject = paragraphPage;
                owners.Add(reference.MarkedContentId, index);
            }
            if (owners.Count != section.ChildElementObjectNumbers.Count ||
                pageObject is not int pageObjectNumber ||
                document.GetPageNumberForObject(pageObjectNumber, budget.CancellationToken) is not int pageNumber) continue;

            // A second candidate story on the same page leaves the order between them ambiguous.
            stories ??= new Dictionary<int, Dictionary<int, int>?>();
            stories[pageNumber] = stories.ContainsKey(pageNumber) ? null : owners;
        }
        return stories;
    }

    private static List<PdfUnderstandingSemanticElement>? BuildTaggedParagraphStory(
        PdfReadPage readPage,
        PdfUnderstandingPageResult page,
        Dictionary<int, int> paragraphOwners,
        int maximum,
        PdfUnderstandingWorkBudget budget) {
        if (paragraphOwners.Count > maximum || page.TextDecodingDiagnostics.Count != 0) return null;
        var ownersByRun = new Dictionary<PdfTextSpan, int>();
        var pageStreams = new HashSet<int?>();
        bool crossesColumns = false;
        int previousOwner = -1;
        PdfTextSpan? previousRun = null;
        foreach (PdfTextSpan run in page.DecodedRuns) {
            budget.Consume(1L + run.Text.Length);
            if (string.IsNullOrWhiteSpace(run.Text)) continue;
            if (!run.MarkedContentId.HasValue ||
                !paragraphOwners.TryGetValue(run.MarkedContentId.Value, out int owner) ||
                (pageStreams.Add(run.ContentStreamObjectNumber) &&
                    !readPage.IsPageContentStreamObjectNumber(run.ContentStreamObjectNumber)) ||
                owner < previousOwner ||
                Math.Abs(run.RotationDegrees) > 0.5D) return null;

            if (previousRun is not null && previousOwner == owner) {
                double em = Math.Max(previousRun.FontSize, run.FontSize);
                crossesColumns |= run.X - previousRun.X > 4D * em && run.Y - previousRun.Y > 3D * em;
            }
            ownersByRun[run] = owner;
            previousOwner = owner;
            previousRun = run;
        }
        if (!crossesColumns) return null;

        var linesByParagraph = new List<PdfUnderstandingLine>?[paragraphOwners.Count];
        var retainedRuns = new HashSet<PdfTextSpan>();
        var retainedWords = new HashSet<PdfUnderstandingWord>();
        foreach (PdfUnderstandingLine line in page.LogicalProjectionLines) {
            budget.Consume();
            if (line.SourceKind != PdfLogicalContentSourceKind.Native) return null;
            var wordsByParagraph = new Dictionary<int, List<PdfUnderstandingWord>>();
            foreach (PdfUnderstandingWord word in line.Words) {
                budget.Consume();
                int owner = -1;
                foreach (PdfTextSpan run in word.SourceRuns) {
                    budget.Consume();
                    if (!ownersByRun.TryGetValue(run, out int runOwner)) {
                        if (!string.IsNullOrWhiteSpace(run.Text)) return null;
                        continue;
                    }
                    if (owner >= 0 && owner != runOwner) return null;
                    owner = runOwner;
                    retainedRuns.Add(run);
                }
                if (owner < 0 || !word.SourceSequence.HasValue || !retainedWords.Add(word)) return null;
                if (!wordsByParagraph.TryGetValue(owner, out var words)) {
                    words = new List<PdfUnderstandingWord>();
                    wordsByParagraph.Add(owner, words);
                }
                words.Add(word);
            }
            foreach (var pair in wordsByParagraph) {
                List<PdfUnderstandingLine> lines = linesByParagraph[pair.Key] ??= new List<PdfUnderstandingLine>();
                lines.Add(PdfAdvancedUnderstandingStages.CreateLineSubset(
                    line, pair.Value, PdfReadingDirection.Auto, budget.Consume, budget.ThrowIfCancellationRequested));
            }
        }
        if (retainedRuns.Count != ownersByRun.Count) return null;

        var evidence = new[] {
            new PdfInferenceEvidence("semantic-tagged-paragraph-story",
                "Reachable tagged paragraphs own the page text in source order across columns.", 0.98D)
        };
        var result = new List<PdfUnderstandingSemanticElement>(linesByParagraph.Length);
        foreach (List<PdfUnderstandingLine>? lines in linesByParagraph) {
            budget.Consume();
            if (lines is null) return null;
            for (int size = lines.Count; size > 0; size /= 2) budget.Consume(lines.Count);
            PdfUnderstandingLine[] ordered = lines.OrderBy(static line => line.SourceSequence).ToArray();
            var region = new PdfUnderstandingRegion(Array.AsReadOnly(ordered), 0.98D, evidence);
            result.Add(new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Paragraph, 0.98D, evidence));
        }
        return result;
    }
}
