namespace OfficeIMO.Pdf;

/// <summary>Defers a canvas's paragraph-relative vertical origin until its first line is laid out.</summary>
internal sealed class PdfCanvasParagraphAnchorItem : PdfCanvasItem {
    internal PdfCanvasParagraphAnchorItem(IReadOnlyList<PdfCanvasItem> items) : base(0D, 0D) {
        Items = items;
    }

    internal IReadOnlyList<PdfCanvasItem> Items { get; }
}

/// <summary>Defers a margin-relative canvas's horizontal origin until its anchor page is laid out.</summary>
internal sealed class PdfCanvasMarginAnchorItem : PdfCanvasItem {
    internal PdfCanvasMarginAnchorItem(IReadOnlyList<PdfCanvasItem> items, double authoredLeftMargin) : base(0D, 0D) {
        Items = items;
        AuthoredLeftMargin = authoredLeftMargin;
    }

    internal IReadOnlyList<PdfCanvasItem> Items { get; }
    internal double AuthoredLeftMargin { get; }
}

/// <summary>Paints anchored drawing content below all body text on its final page.</summary>
internal sealed class PdfCanvasBehindTextItem : PdfCanvasItem {
    internal PdfCanvasBehindTextItem(IReadOnlyList<PdfCanvasItem> items, long zOrder) : base(0D, 0D) {
        Items = items;
        ZOrder = zOrder;
    }

    internal IReadOnlyList<PdfCanvasItem> Items { get; }
    internal long ZOrder { get; }
}
