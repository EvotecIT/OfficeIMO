namespace OfficeIMO.Pdf;

public sealed partial class PdfPageCanvas {
    /// <summary>
    /// Groups positioned text fragments under one logical replacement string for extraction and accessibility.
    /// Child paint remains unchanged while readers that honor <c>ActualText</c> receive the supplied logical text once.
    /// </summary>
    public PdfPageCanvas ActualText(string text, Action<PdfPageCanvas> build) {
        return AddActualText(text, 0D, 0D, hasPosition: false, build);
    }

    /// <summary>
    /// Groups positioned paint under one logical replacement string and anchors the invisible
    /// extraction span at an absolute top-left page coordinate.
    /// </summary>
    public PdfPageCanvas ActualText(string text, double x, double y, Action<PdfPageCanvas> build) {
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        return AddActualText(text, x, y, hasPosition: true, build);
    }

    /// <summary>
    /// Groups paint under a logical replacement with an extraction rectangle in
    /// top-left page coordinates. The rectangle preserves spacing between adjacent
    /// spans; the paint remains owned by the same replacement for extraction and editing.
    /// </summary>
    public PdfPageCanvas ActualText(string text, double x, double y, double width, double height, Action<PdfPageCanvas> build) {
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        return AddActualText(text, x, y, hasPosition: true, build, width, height);
    }

    private PdfPageCanvas AddActualText(string text, double x, double y, bool hasPosition, Action<PdfPageCanvas> build,
        double width = 0D, double height = 0D) {
        Guard.NotNull(text, nameof(text));
        if (text.Length == 0) throw new ArgumentException("Canvas actual text cannot be empty.", nameof(text));
        Guard.NotNull(build, nameof(build));
        var nestedCanvas = new PdfPageCanvas(allowOutOfPageCoordinates: true);
        build(nestedCanvas);
        if (nestedCanvas.Items.Count == 0) {
            throw new ArgumentException("Canvas actual-text groups require at least one content item.", nameof(build));
        }
        _items.Add(new PdfCanvasActualTextItem(text, x, y, hasPosition, nestedCanvas.Items, width, height));
        return this;
    }

}

internal sealed class PdfCanvasActualTextItem : PdfCanvasItem {
    public PdfCanvasActualTextItem(string text, double x, double y, bool hasPosition, IReadOnlyList<PdfCanvasItem> items, double width = 0D, double height = 0D)
        : base(x, y) {
        Text = text;
        HasPosition = hasPosition;
        Items = items;
        Width = width;
        Height = height;
    }

    public double Width { get; }
    public double Height { get; }
    public bool UsesBounds => Width > 0D && Height > 0D;
    public string Text { get; }
    public bool HasPosition { get; }
    public IReadOnlyList<PdfCanvasItem> Items { get; }
}
