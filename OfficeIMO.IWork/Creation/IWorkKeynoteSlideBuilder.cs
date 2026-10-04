namespace OfficeIMO.IWork;

/// <summary>A slide being created in an <see cref="IWorkKeynoteDocument"/>.</summary>
/// <remarks>Text boxes retain insertion order, which also determines stacking order. Instances are not thread safe.</remarks>
public sealed class IWorkKeynoteSlideBuilder {
    private readonly IWorkCanvasSize _canvas;
    private readonly List<IWorkKeynoteTextBox> _textBoxes = new();

    internal IWorkKeynoteSlideBuilder(IWorkCanvasSize canvas, IWorkColor backgroundColor) {
        _canvas = canvas;
        BackgroundColor = backgroundColor;
        TextBoxes = _textBoxes.AsReadOnly();
    }

    /// <summary>Gets the opaque sRGB slide background.</summary>
    public IWorkColor BackgroundColor { get; }
    /// <summary>Gets the positioned text boxes in stacking order.</summary>
    public IReadOnlyList<IWorkKeynoteTextBox> TextBoxes { get; }

    /// <summary>Adds a left-aligned text box with zero internal padding and one font and color throughout.</summary>
    /// <remarks>Dimensions are in points and must lie within the canvas. CR and CRLF normalize to LF.
    /// Font availability and text wrapping depend on the native application's fonts; this method does not measure or fit text.</remarks>
    public IWorkKeynoteTextBox AddText(string text, float leftPoints, float topPoints, float widthPoints,
        float heightPoints, string fontName = "Arial", float fontSizePoints = 32, string color = "000000") {
        var box = new IWorkKeynoteTextBox(text, leftPoints, topPoints, widthPoints, heightPoints,
            fontName, fontSizePoints, IWorkKeynoteTextBox.ParseColor(color));
        if (leftPoints < 0 || topPoints < 0 || (double)leftPoints + widthPoints > _canvas.WidthPoints ||
            (double)topPoints + heightPoints > _canvas.HeightPoints) {
            throw new ArgumentOutOfRangeException(nameof(leftPoints), "The text box must lie within the slide canvas.");
        }
        _textBoxes.Add(box);
        return box;
    }
}
