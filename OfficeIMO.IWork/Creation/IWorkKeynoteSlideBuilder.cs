namespace OfficeIMO.IWork;

/// <summary>A slide being created in an <see cref="IWorkKeynoteDocument"/>.</summary>
/// <remarks>Text boxes, images and tables retain their combined insertion order, which determines stacking order. Instances are not thread safe.</remarks>
public sealed class IWorkKeynoteSlideBuilder {
    private readonly IWorkCanvasSize _canvas;
    private readonly List<IWorkKeynoteTextBox> _textBoxes = new();
    private readonly List<IWorkKeynoteImage> _images = new();
    private readonly List<IWorkKeynoteTable> _tables = new();
    private readonly List<IWorkKeynoteCreationDrawable> _drawables = new();

    internal IWorkKeynoteSlideBuilder(IWorkCanvasSize canvas, IWorkColor backgroundColor) {
        _canvas = canvas;
        BackgroundColor = backgroundColor;
        TextBoxes = _textBoxes.AsReadOnly();
        Images = _images.AsReadOnly();
        Tables = _tables.AsReadOnly();
    }

    /// <summary>Gets the opaque sRGB slide background.</summary>
    public IWorkColor BackgroundColor { get; }
    /// <summary>Gets the positioned text boxes in stacking order.</summary>
    public IReadOnlyList<IWorkKeynoteTextBox> TextBoxes { get; }
    /// <summary>Gets the positioned images in their relative stacking order.</summary>
    public IReadOnlyList<IWorkKeynoteImage> Images { get; }
    /// <summary>Gets the positioned tables in their relative stacking order.</summary>
    public IReadOnlyList<IWorkKeynoteTable> Tables { get; }
    internal IReadOnlyList<IWorkKeynoteCreationDrawable> Drawables => _drawables;

    /// <summary>Adds a snapshot of scalar table data with equally sized rows and columns.</summary>
    /// <remarks>Dimensions are in points and must lie within the canvas. The visible table title is disabled.
    /// All cells share white fill, four-point padding and the selected font and text color. Content is not measured to fit.</remarks>
    public IWorkKeynoteTable AddTable(IWorkTableData data, float leftPoints, float topPoints, float widthPoints,
        float heightPoints, string name = "Table", string fontName = "Arial", float fontSizePoints = 20, string color = "000000") {
        ValidateBounds(leftPoints, topPoints, widthPoints, heightPoints);
        var table = new IWorkKeynoteTable(data, leftPoints, topPoints, widthPoints, heightPoints,
            name, fontName, fontSizePoints, IWorkCreationValidation.ParseColor(color));
        _tables.Add(table); _drawables.Add(table);
        return table;
    }

    /// <summary>Adds a complete PNG or JPEG and copies its payload into an immutable image.</summary>
    /// <remarks>Dimensions are in points and must lie within the canvas. The full image fills the frame,
    /// including stretching when its aspect ratio differs. No crop, mask, rotation or image adjustment is applied.
    /// Core image limits apply during validation; aggregate write limits apply when saving.</remarks>
    public IWorkKeynoteImage AddImage(byte[] data, float leftPoints, float topPoints, float widthPoints,
        float heightPoints, string? accessibilityDescription = null, System.Threading.CancellationToken cancellationToken = default) {
        ValidateBounds(leftPoints, topPoints, widthPoints, heightPoints);
        var image = new IWorkKeynoteImage(data, leftPoints, topPoints, widthPoints, heightPoints,
            accessibilityDescription, cancellationToken);
        _images.Add(image); _drawables.Add(image);
        return image;
    }

    /// <summary>Adds a left-aligned text box with zero internal padding and one font and color throughout.</summary>
    /// <remarks>Dimensions are in points and must lie within the canvas. CR and CRLF normalize to LF.
    /// Font availability and text wrapping depend on the native application's fonts; this method does not measure or fit text.</remarks>
    public IWorkKeynoteTextBox AddText(string text, float leftPoints, float topPoints, float widthPoints,
        float heightPoints, string fontName = "Arial", float fontSizePoints = 32, string color = "000000") {
        var box = new IWorkKeynoteTextBox(text, leftPoints, topPoints, widthPoints, heightPoints,
            fontName, fontSizePoints, IWorkCreationValidation.ParseColor(color));
        return AddBox(box, leftPoints, topPoints, widthPoints, heightPoints);
    }

    /// <summary>Adds a left-aligned rich-text box with zero padding and immutable character runs.</summary>
    /// <remarks>Runs are copied when added. Empty runs have no visible range and are discarded unless the entire box is blank.
    /// Dimensions are in points and must lie within the canvas. This method does not measure or fit text.</remarks>
    public IWorkKeynoteTextBox AddRichText(IEnumerable<IWorkTextRunData> runs, float leftPoints, float topPoints,
        float widthPoints, float heightPoints) {
        var box = new IWorkKeynoteTextBox(runs, leftPoints, topPoints, widthPoints, heightPoints);
        return AddBox(box, leftPoints, topPoints, widthPoints, heightPoints);
    }

    private IWorkKeynoteTextBox AddBox(IWorkKeynoteTextBox box, float leftPoints, float topPoints,
        float widthPoints, float heightPoints) {
        ValidateBounds(leftPoints, topPoints, widthPoints, heightPoints);
        _textBoxes.Add(box); _drawables.Add(box);
        return box;
    }

    private void ValidateBounds(float leftPoints, float topPoints, float widthPoints, float heightPoints) {
        if (leftPoints < 0 || topPoints < 0 || (double)leftPoints + widthPoints > _canvas.WidthPoints ||
            (double)topPoints + heightPoints > _canvas.HeightPoints) {
            throw new ArgumentOutOfRangeException(nameof(leftPoints), "The drawable must lie within the slide canvas.");
        }
    }
}
