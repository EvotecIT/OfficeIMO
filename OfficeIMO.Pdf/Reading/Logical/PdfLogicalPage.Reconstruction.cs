using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    internal static PdfLogicalPage From(
        PdfReadDocument document,
        PdfReadPage page,
        int pageNumber,
        PdfTextLayoutOptions? options,
        IReadOnlyList<PdfLogicalFormWidget>? pageFormWidgets = null,
        PdfUnderstandingPageResult? analysis = null,
        bool includeImages = true,
        bool includeFormWidgets = true,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var size = page.GetPageSize();
        PdfPageGeometry geometry = page.GetGeometry();
        List<PdfUnderstandingLine>? retainedLines = analysis is null
            ? null
            : GetRetainedProjectionLines(analysis);
        IReadOnlyList<PdfTextSpan>? retainedRuns = analysis is null
            ? null
            : analysis.RestrictLogicalProjectionToReadingOrder
                ? GetRetainedProjectionRuns(retainedLines!, cancellationToken)
                : analysis.DecodedRuns;
        IReadOnlyList<PdfUnderstandingTableCandidate>? projectedTableCandidates = analysis is null
            ? null
            : GetProjectedTableCandidates(analysis, retainedLines!, retainedRuns!, cancellationToken);
        var structured = analysis is null
            ? page.ExtractStructured(options)
            : CreateStructuredProjection(
                page,
                retainedLines!,
                projectedTableCandidates!,
                analysis,
                options,
                cancellationToken);
        var elements = new List<IPdfLogicalElement>();
        var textBlocks = new List<PdfLogicalTextBlock>();
        var semanticByTextBlock = new Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement>();
        var tables = new List<PdfLogicalTable>();
        var images = new List<PdfLogicalImage>();
        var links = new List<PdfLogicalLinkAnnotation>();
        var formWidgets = new List<PdfLogicalFormWidget>();
        var listLines = new HashSet<string>(structured.ListItems.Select(NormalizeForKindComparison), StringComparer.Ordinal);
        PdfUnderstandingPageResult pageAnalysis = analysis ?? PdfUnderstandingPageResult.Empty(pageNumber);
        SemanticElementIndex semanticIndex = SemanticElementIndex.Create(pageAnalysis.Elements);
        var headingSourceRuns = new HashSet<PdfTextSpan>(
            structured.Headings.SelectMany(static heading => heading.Line.Spans));
        LogicalTableSourceIndex tableSourceIndex = LogicalTableSourceIndex.Create(
            analysis is null ? structured.TablesDetailed : null,
            projectedTableCandidates);

        foreach (var line in structured.LinesDetailed) {
            cancellationToken.ThrowIfCancellationRequested();
            string text = line.Text?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            bool isStructuredHeading = analysis is null && line.Spans.Any(headingSourceRuns.Contains);
            bool isStructuredListItem = analysis is null &&
                (listLines.Contains(NormalizeForKindComparison(text)) || ContentStructureExtractor.IsListItemText(text));
            PdfUnderstandingSemanticElement? semantic = semanticIndex.Find(line.Y, line.XStart, text, line.Spans);
            var kind = ToLogicalKind(semantic, line, isStructuredHeading)
                ?? (isStructuredHeading
                    ? PdfLogicalElementKind.Heading
                    : isStructuredListItem
                    ? PdfLogicalElementKind.ListItem
                    : PdfLogicalElementKind.TextBlock);
            var block = new PdfLogicalTextBlock(
                pageNumber,
                kind,
                text,
                line.XStart,
                line.XEnd,
                line.Y,
                line.FontSize,
                line.Spans,
                line.SourceKind,
                line.Confidence,
                visualBounds: line.VisualBounds,
                isTableContent: tableSourceIndex.Contains(line));
            textBlocks.Add(block);
            if (semantic is not null) semanticByTextBlock.Add(block, semantic);
            elements.Add(block);
        }

        foreach (var row in structured.LeaderRows) {
            if (row.Length < 2) {
                continue;
            }

            var leader = new PdfLogicalLeaderRow(pageNumber, row[0], row[1]);
            elements.Add(leader);
        }

        IEnumerable<PdfLogicalTable> logicalTables = analysis is null
            ? structured.TablesDetailed.Select(table => PdfLogicalTable.From(pageNumber, table))
            : projectedTableCandidates!.Select(table => PdfLogicalTable.From(pageNumber, table));
        foreach (PdfLogicalTable logicalTable in logicalTables) {
            tables.Add(logicalTable);
            elements.Add(logicalTable);
        }

        IReadOnlyList<PdfImagePlacement> imagePlacements = !includeImages ? Array.Empty<PdfImagePlacement>() : analysis is not null
            ? pageAnalysis.ImagePlacements
            : page.GetImagePlacements(pageNumber, cancellationToken);
        foreach (var image in includeImages ? page.GetImages(pageNumber, imagePlacements, cancellationToken) : Array.Empty<PdfExtractedImage>()) {
            IReadOnlyList<PdfImagePlacement> matchingPlacements = MatchImagePlacements(image, imagePlacements);
            var logicalImage = new PdfLogicalImage(
                image,
                matchingPlacements,
                MatchImageRegions(matchingPlacements, pageAnalysis.ImageRegions));
            images.Add(logicalImage);
            elements.Add(logicalImage);
        }

        IReadOnlyList<PdfLinkAnnotation> readLinkAnnotations = page.GetLinkAnnotations();
        var linkAnnotations = new List<PdfLinkAnnotation>(readLinkAnnotations.Count);
        for (int i = 0; i < readLinkAnnotations.Count; i++) {
            PdfLinkAnnotation linkAnnotation = ResolveLinkDestinationPageNumber(document, readLinkAnnotations[i], cancellationToken);
            linkAnnotations.Add(linkAnnotation);
            var logicalLink = new PdfLogicalLinkAnnotation(pageNumber, linkAnnotation);
            links.Add(logicalLink);
            elements.Add(logicalLink);
        }

        IReadOnlyList<PdfAnnotation> readAnnotations = page.GetAnnotations();
        var annotations = new List<PdfAnnotation>(readAnnotations.Count);
        for (int i = 0; i < readAnnotations.Count; i++) {
            annotations.Add(readAnnotations[i].WithPageNumber(pageNumber));
        }

        if (pageFormWidgets is not null) {
            for (int widgetIndex = 0; widgetIndex < pageFormWidgets.Count; widgetIndex++) {
                PdfLogicalFormWidget logicalWidget = pageFormWidgets[widgetIndex];
                formWidgets.Add(logicalWidget);
                elements.Add(logicalWidget);
            }
        }

        IReadOnlyList<PdfPageAction> readPageActions = page.GetPageActions();
        var pageActions = new List<PdfPageAction>(readPageActions.Count);
        for (int i = 0; i < readPageActions.Count; i++) {
            pageActions.Add(readPageActions[i].WithPageNumber(pageNumber));
        }

        (int vectorPrimitiveCount, int unrepresentedVectorPrimitiveCount) =
            includeImages ? page.GetVisibleVisualPrimitiveCounts(structured.TablesDetailed) : (0, 0);
        Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>> textBlockLookup =
            CreateLogicalTextBlockLookup(textBlocks);
        LogicalTextBlockSourceIndex textBlockSourceIndex = LogicalTextBlockSourceIndex.Create(textBlocks);
        return new PdfLogicalPage(
            pageNumber,
            size.Width,
            size.Height,
            page.GetRotationDegrees(),
            geometry,
            elements.AsReadOnly(),
            textBlocks.AsReadOnly(),
            BuildHeadings(pageNumber, analysis is null ? structured.Headings : new List<StructuredHeading>(), textBlocks, semanticByTextBlock, semanticIndex, textBlockLookup, textBlockSourceIndex),
            analysis is null
                ? BuildParagraphs(pageNumber, structured.Paragraphs, textBlocks, textBlockSourceIndex)
                : BuildParagraphsFromSemanticRegions(pageNumber, textBlocks, semanticByTextBlock),
            BuildListItems(pageNumber, analysis is null ? structured.ListNodes : new List<StructuredListItem>(), textBlocks, semanticByTextBlock, textBlockLookup, textBlockSourceIndex, pageAnalysis),
            tables.AsReadOnly(),
            vectorPrimitiveCount,
            unrepresentedVectorPrimitiveCount,
            images.AsReadOnly(),
            links.AsReadOnly(),
            annotations.AsReadOnly(),
            linkAnnotations.AsReadOnly(),
            formWidgets.AsReadOnly(),
            pageActions.AsReadOnly(),
            page.HasOptionalContentUsage(cancellationToken),
            pageAnalysis, hasVisualPrimitiveCounts: includeImages,
            includesImages: includeImages, includesFormWidgets: includeFormWidgets);
    }

}
