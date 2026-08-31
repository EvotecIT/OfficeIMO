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
        return new PdfLogicalPage(
            PageNumber,
            Width,
            Height,
            RotationDegrees,
            Geometry,
            elements.AsReadOnly(),
            TextBlocks.Concat(ocrTextBlocks).ToArray(),
            Headings.Concat(ocrHeadings).ToArray(),
            Paragraphs.Concat(ocrParagraphs).ToArray(),
            ListItems.Concat(ocrListItems).ToArray(),
            Tables.Concat(ocrTables).ToArray(),
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
}
