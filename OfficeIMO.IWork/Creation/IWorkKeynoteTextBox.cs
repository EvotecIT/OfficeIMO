namespace OfficeIMO.IWork;

/// <summary>An immutable positioned text box in a newly created Keynote presentation.</summary>
public sealed class IWorkKeynoteTextBox : IWorkKeynoteCreationDrawable {
    internal IWorkKeynoteTextBox(string text, float left, float top, float width, float height,
        string fontName, float fontSize, IWorkColor color)
        : this(new[] { new IWorkKeynoteTextRun(text, fontName, fontSize, color) }, left, top, width, height) { }

    internal IWorkKeynoteTextBox(IEnumerable<IWorkKeynoteTextRun> runs, float left, float top, float width, float height) {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        IWorkCreationValidation.ValidateFinite(left, nameof(left));
        IWorkCreationValidation.ValidateFinite(top, nameof(top));
        IWorkCreationValidation.ValidatePositive(width, nameof(width));
        IWorkCreationValidation.ValidatePositive(height, nameof(height));
        IWorkKeynoteTextRun[] input = runs.ToArray();
        if (input.Length == 0 || input.Any(run => run == null))
            throw new ArgumentException("At least one non-null character run is required.", nameof(runs));
        // Empty ranges have no native character position. Keep one default run
        // for a blank box, otherwise discard them without emitting duplicate offsets.
        IWorkKeynoteTextRun[] retained = input.Where(run => run.Text.Length > 0).ToArray();
        if (retained.Length == 0) retained = new[] { input[0] };
        Runs = Array.AsReadOnly(retained);
        Text = string.Concat(retained.Select(run => run.Text));
        Geometry = new IWorkGeometry(left, top, width, height, 0);
        FontName = retained[0].FontName;
        FontSizePoints = retained[0].FontSizePoints;
        Color = retained[0].Color;
    }

    /// <summary>Gets the text with normalized LF line breaks.</summary>
    public string Text { get; }
    /// <summary>Gets the immutable character runs in text order.</summary>
    public IReadOnlyList<IWorkKeynoteTextRun> Runs { get; }
    /// <summary>Gets the frame geometry in points.</summary>
    public IWorkGeometry Geometry { get; }
    /// <summary>Gets the first run's font, also used as the paragraph default.</summary>
    public string FontName { get; }
    /// <summary>Gets the first run's font size, also used as the paragraph default.</summary>
    public float FontSizePoints { get; }
    /// <summary>Gets the first run's opaque sRGB color, also used as the paragraph default.</summary>
    public IWorkColor Color { get; }

}
