using System.Globalization;
using System.Text;

namespace OfficeIMO.IWork;

/// <summary>Shared scalar validation for template-free iWork creation models.</summary>
internal static class IWorkCreationValidation {
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal static void ValidateFontName(string fontName) {
        if (fontName == null) throw new ArgumentNullException(nameof(fontName));
        if (string.IsNullOrWhiteSpace(fontName) || fontName.Length > 128 || fontName.Any(char.IsControl))
            throw new ArgumentException("A font name of at most 128 characters without control characters is required.", nameof(fontName));
        StrictUtf8.GetByteCount(fontName);
    }

    internal static string NormalizeText(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Any(character => char.IsControl(character) && character != '\n'))
            throw new ArgumentException("Text supports line breaks but no other control characters.", nameof(text));
        // Invalid UTF-16 must not silently change to replacement glyphs when encoded.
        StrictUtf8.GetByteCount(normalized);
        return normalized;
    }

    internal static void ValidateFinite(float value, string name) {
        if (float.IsNaN(value) || float.IsInfinity(value)) throw new ArgumentOutOfRangeException(name, "A finite value is required.");
    }

    internal static void ValidatePositive(float value, string name) {
        ValidateFinite(value, name);
        if (value <= 0) throw new ArgumentOutOfRangeException(name, "A positive value is required.");
    }

    internal static IWorkColor ParseColor(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Length != 6 || !uint.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out uint rgb))
            throw new ArgumentException("An opaque six-digit sRGB color is required, for example E6F2FF.", nameof(value));
        return new IWorkColor((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb, 255);
    }
}
