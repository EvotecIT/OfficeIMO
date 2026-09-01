namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    internal PdfLogicalPage WithOcrContent(
        IReadOnlyList<PdfLogicalTextBlock> ocrTextBlocks,
        IReadOnlyList<PdfLogicalHeading> ocrHeadings,
        IReadOnlyList<PdfLogicalParagraph> ocrParagraphs,
        IReadOnlyList<PdfLogicalListItem> ocrListItems,
        IReadOnlyList<PdfLogicalTable> ocrTables) {
        var elements = new List<IPdfLogicalElement>(Elements.Count + ocrTextBlocks.Count + ocrTables.Count);
        elements.AddRange(Elements);
        elements.AddRange(ocrTextBlocks);
        elements.AddRange(ocrTables);
        PdfLogicalTextBlock[] textBlocks = TextBlocks.Concat(ocrTextBlocks).ToArray();
        PdfLogicalHeading[] headings = Headings.Concat(ocrHeadings).ToArray();
        PdfLogicalParagraph[] paragraphs = Paragraphs.Concat(ocrParagraphs).ToArray();
        PdfLogicalListItem[] listItems = ListItems.Concat(ocrListItems).ToArray();
        PdfLogicalTable[] tables = Tables.Concat(ocrTables).ToArray();
        var merged = new PdfLogicalPage(
            PageNumber,
            Width,
            Height,
            RotationDegrees,
            Geometry,
            elements.AsReadOnly(),
            textBlocks,
            headings,
            paragraphs,
            listItems,
            tables,
            VectorPrimitiveCount,
            UnrepresentedVectorPrimitiveCount,
            Images,
            Links,
            Annotations,
            LinkAnnotations,
            FormWidgets,
            PageActions,
            Analysis);
        IReadOnlyList<PdfLogicalTextBlock> orderedTextBlocks = OrderTextBlocks(merged);
        if (orderedTextBlocks.SequenceEqual(textBlocks)) return merged;
        var orderedElements = new List<IPdfLogicalElement>(elements.Count);
        orderedElements.AddRange(orderedTextBlocks);
        orderedElements.AddRange(elements.Where(static element => element is not PdfLogicalTextBlock));
        return new PdfLogicalPage(
            PageNumber,
            Width,
            Height,
            RotationDegrees,
            Geometry,
            orderedElements.AsReadOnly(),
            orderedTextBlocks,
            headings,
            paragraphs,
            listItems,
            tables,
            VectorPrimitiveCount,
            UnrepresentedVectorPrimitiveCount,
            Images,
            Links,
            Annotations,
            LinkAnnotations,
            FormWidgets,
            PageActions,
            Analysis);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfLogicalTextBlock> OrderTextBlocks(PdfLogicalPage page) {
        var available = new HashSet<PdfLogicalTextBlock>(page.TextBlocks);
        var seen = new HashSet<PdfLogicalTextBlock>();
        var ordered = new List<PdfLogicalTextBlock>(page.TextBlocks.Count);
        foreach (PdfLogicalTextBlock header in page.TextBlocks
                     .Where(static block => block.Kind == PdfLogicalElementKind.Header)
                     .OrderBy(block => GetVisualPosition(page, block).Top)
                     .ThenBy(block => GetVisualPosition(page, block).Left)) {
            Add(header);
        }
        IReadOnlyList<PdfLogicalReadingOrderItem> readingOrder = PdfLogicalReadingOrderAnalysis.Analyze(page);
        for (int itemIndex = 0; itemIndex < readingOrder.Count; itemIndex++) {
            PdfLogicalReadingOrderItem item = readingOrder[itemIndex];
            switch (item.Kind) {
                case PdfLogicalReadingOrderKind.TextBlock:
                    Add(page.TextBlocks[item.SourceIndex]);
                    break;
                case PdfLogicalReadingOrderKind.Heading:
                    Add(page.Headings[item.SourceIndex].Line);
                    break;
                case PdfLogicalReadingOrderKind.Paragraph:
                    foreach (PdfLogicalTextBlock line in page.Paragraphs[item.SourceIndex].Lines) Add(line);
                    break;
                case PdfLogicalReadingOrderKind.ListItem:
                    foreach (PdfLogicalTextBlock line in page.ListItems[item.SourceIndex].Lines) Add(line);
                    break;
            }
        }
        foreach (PdfLogicalTextBlock block in page.TextBlocks.Where(static block => block.Kind != PdfLogicalElementKind.Footer)) Add(block);
        foreach (PdfLogicalTextBlock footer in page.TextBlocks
                     .Where(static block => block.Kind == PdfLogicalElementKind.Footer)
                     .OrderBy(block => GetVisualPosition(page, block).Top)
                     .ThenBy(block => GetVisualPosition(page, block).Left)) {
            Add(footer);
        }
        return ordered.AsReadOnly();

        void Add(PdfLogicalTextBlock block) {
            if (available.Contains(block) && seen.Add(block)) ordered.Add(block);
        }
    }

    private static (double Top, double Left) GetVisualPosition(PdfLogicalPage page, PdfLogicalTextBlock block) {
        if (block.VisualBounds is PdfLogicalVisualBounds visual) return (visual.Top, visual.Left);
        PdfVisualBounds transformed = page.TransformBoundsToVisual(
            block.XStart,
            block.BaselineY,
            block.XEnd,
            block.BaselineY + Math.Max(1D, block.FontSize));
        return (transformed.Top, transformed.Left);
    }
}
