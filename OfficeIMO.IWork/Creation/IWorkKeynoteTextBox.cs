using System.Globalization;
using System.Text;

namespace OfficeIMO.IWork;

/// <summary>An immutable positioned text box in a newly created Keynote presentation.</summary>
public sealed class IWorkKeynoteTextBox {
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal IWorkKeynoteTextBox(string text, float left, float top, float width, float height,
        string fontName, float fontSize, IWorkColor color) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (fontName == null) throw new ArgumentNullException(nameof(fontName));
        ValidateFinite(left, nameof(left));
        ValidateFinite(top, nameof(top));
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        ValidatePositive(fontSize, nameof(fontSize));
        if (string.IsNullOrWhiteSpace(fontName) || fontName.Length > 128 || fontName.Any(char.IsControl)) {
            throw new ArgumentException("A font name of at most 128 characters without control characters is required.", nameof(fontName));
        }
        Text = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (Text.Any(character => char.IsControl(character) && character != '\n')) {
            throw new ArgumentException("Text supports line breaks but no other control characters.", nameof(text));
        }
        // Do not silently replace unpaired UTF-16 surrogates with a different glyph.
        StrictUtf8.GetByteCount(Text);
        StrictUtf8.GetByteCount(fontName);
        Geometry = new IWorkGeometry(left, top, width, height, 0);
        FontName = fontName;
        FontSizePoints = fontSize;
        Color = color;
    }

    /// <summary>Gets the text with normalized LF line breaks.</summary>
    public string Text { get; }
    /// <summary>Gets the frame geometry in points.</summary>
    public IWorkGeometry Geometry { get; }
    /// <summary>Gets the requested native font name.</summary>
    public string FontName { get; }
    /// <summary>Gets the font size in points.</summary>
    public float FontSizePoints { get; }
    /// <summary>Gets the opaque sRGB text color.</summary>
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
