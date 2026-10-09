using System.Globalization;
using System.Text;

namespace OfficeIMO.IWork;

/// <summary>An immutable positioned text box in a newly created Keynote presentation.</summary>
public sealed class IWorkKeynoteTextBox {
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal IWorkKeynoteTextBox(string text, float left, float top, float width, float height,
        string fontName, float fontSize, IWorkColor color)
        : this(new[] { new IWorkKeynoteTextRun(text, fontName, fontSize, color) }, left, top, width, height) { }

    internal IWorkKeynoteTextBox(IEnumerable<IWorkKeynoteTextRun> runs, float left, float top, float width, float height) {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        ValidateFinite(left, nameof(left));
        ValidateFinite(top, nameof(top));
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
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

    internal static void ValidateFontName(string fontName) {
        if (fontName == null) throw new ArgumentNullException(nameof(fontName));
        if (string.IsNullOrWhiteSpace(fontName) || fontName.Length > 128 || fontName.Any(char.IsControl)) {
            throw new ArgumentException("A font name of at most 128 characters without control characters is required.", nameof(fontName));
        }
        StrictUtf8.GetByteCount(fontName);
    }

    internal static string NormalizeText(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Any(character => char.IsControl(character) && character != '\n')) {
            throw new ArgumentException("Text supports line breaks but no other control characters.", nameof(text));
        }
        // Do not silently replace unpaired UTF-16 surrogates with a different glyph.
        StrictUtf8.GetByteCount(normalized);
        return normalized;
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

    internal static void ValidateFinite(float value, string name) {
        if (float.IsNaN(value) || float.IsInfinity(value)) throw new ArgumentOutOfRangeException(name, "A finite value is required.");
    }

    internal static void ValidatePositive(float value, string name) {
        ValidateFinite(value, name);
        if (value <= 0) throw new ArgumentOutOfRangeException(name, "A positive value is required.");
    }

    internal static IWorkColor ParseColor(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Length != 6 || !uint.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out uint rgb)) {
            throw new ArgumentException("An opaque six-digit sRGB color is required, for example E6F2FF.", nameof(value));
        }
        return new IWorkColor((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb, 255);
    }
}
