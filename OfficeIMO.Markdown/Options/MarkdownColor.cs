using System.Globalization;

namespace OfficeIMO.Markdown;

/// <summary>
/// Renderer-neutral RGB color used by shared Markdown visual themes.
/// </summary>
public readonly struct MarkdownColor : IEquatable<MarkdownColor> {
    private static readonly Dictionary<string, MarkdownColor> NamedColors = new Dictionary<string, MarkdownColor>(StringComparer.OrdinalIgnoreCase) {
        ["Transparent"] = new MarkdownColor(0, 0, 0, 0),
        ["AliceBlue"] = FromRgb(240, 248, 255),
        ["Aqua"] = FromRgb(0, 255, 255),
        ["Aquamarine"] = FromRgb(127, 255, 212),
        ["Azure"] = FromRgb(240, 255, 255),
        ["Beige"] = FromRgb(245, 245, 220),
        ["Black"] = FromRgb(0, 0, 0),
        ["Blue"] = FromRgb(0, 0, 255),
        ["BlueViolet"] = FromRgb(138, 43, 226),
        ["Brown"] = FromRgb(165, 42, 42),
        ["Coral"] = FromRgb(255, 127, 80),
        ["CornflowerBlue"] = FromRgb(100, 149, 237),
        ["Crimson"] = FromRgb(220, 20, 60),
        ["Cyan"] = FromRgb(0, 255, 255),
        ["DarkBlue"] = FromRgb(0, 0, 139),
        ["DarkCyan"] = FromRgb(0, 139, 139),
        ["DarkGoldenrod"] = FromRgb(184, 134, 11),
        ["DarkGray"] = FromRgb(169, 169, 169),
        ["DarkGrey"] = FromRgb(169, 169, 169),
        ["DarkGreen"] = FromRgb(0, 100, 0),
        ["DarkKhaki"] = FromRgb(189, 183, 107),
        ["DarkMagenta"] = FromRgb(139, 0, 139),
        ["DarkOliveGreen"] = FromRgb(85, 107, 47),
        ["DarkOrange"] = FromRgb(255, 140, 0),
        ["DarkRed"] = FromRgb(139, 0, 0),
        ["DarkSalmon"] = FromRgb(233, 150, 122),
        ["DarkSeaGreen"] = FromRgb(143, 188, 143),
        ["DarkSlateBlue"] = FromRgb(72, 61, 139),
        ["DeepPink"] = FromRgb(255, 20, 147),
        ["DimGray"] = FromRgb(105, 105, 105),
        ["DimGrey"] = FromRgb(105, 105, 105),
        ["DodgerBlue"] = FromRgb(30, 144, 255),
        ["ForestGreen"] = FromRgb(34, 139, 34),
        ["Fuchsia"] = FromRgb(255, 0, 255),
        ["Gold"] = FromRgb(255, 215, 0),
        ["Goldenrod"] = FromRgb(218, 165, 32),
        ["Gray"] = FromRgb(128, 128, 128),
        ["Grey"] = FromRgb(128, 128, 128),
        ["Green"] = FromRgb(0, 128, 0),
        ["GreenYellow"] = FromRgb(173, 255, 47),
        ["Honeydew"] = FromRgb(240, 255, 240),
        ["HotPink"] = FromRgb(255, 105, 180),
        ["IndianRed"] = FromRgb(205, 92, 92),
        ["Indigo"] = FromRgb(75, 0, 130),
        ["Khaki"] = FromRgb(240, 230, 140),
        ["Lavender"] = FromRgb(230, 230, 250),
        ["LightBlue"] = FromRgb(173, 216, 230),
        ["LightCoral"] = FromRgb(240, 128, 128),
        ["LightCyan"] = FromRgb(224, 255, 255),
        ["LightGoldenrodYellow"] = FromRgb(250, 250, 210),
        ["LightGray"] = FromRgb(211, 211, 211),
        ["LightGrey"] = FromRgb(211, 211, 211),
        ["LightGreen"] = FromRgb(144, 238, 144),
        ["LightPink"] = FromRgb(255, 182, 193),
        ["LightSkyBlue"] = FromRgb(135, 206, 250),
        ["LightSteelBlue"] = FromRgb(176, 196, 222),
        ["LightYellow"] = FromRgb(255, 255, 224),
        ["Lime"] = FromRgb(0, 255, 0),
        ["Magenta"] = FromRgb(255, 0, 255),
        ["Maroon"] = FromRgb(128, 0, 0),
        ["MediumAquamarine"] = FromRgb(102, 205, 170),
        ["MediumPurple"] = FromRgb(147, 112, 219),
        ["MediumSeaGreen"] = FromRgb(60, 179, 113),
        ["MediumVioletRed"] = FromRgb(199, 21, 133),
        ["MistyRose"] = FromRgb(255, 228, 225),
        ["Navy"] = FromRgb(0, 0, 128),
        ["Olive"] = FromRgb(128, 128, 0),
        ["Orange"] = FromRgb(255, 165, 0),
        ["OrangeRed"] = FromRgb(255, 69, 0),
        ["Peru"] = FromRgb(205, 133, 63),
        ["Pink"] = FromRgb(255, 192, 203),
        ["Plum"] = FromRgb(221, 160, 221),
        ["Purple"] = FromRgb(128, 0, 128),
        ["RebeccaPurple"] = FromRgb(102, 51, 153),
        ["Red"] = FromRgb(255, 0, 0),
        ["SaddleBrown"] = FromRgb(139, 69, 19),
        ["Salmon"] = FromRgb(250, 128, 114),
        ["SandyBrown"] = FromRgb(244, 164, 96),
        ["SeaGreen"] = FromRgb(46, 139, 87),
        ["Silver"] = FromRgb(192, 192, 192),
        ["SkyBlue"] = FromRgb(135, 206, 235),
        ["SteelBlue"] = FromRgb(70, 130, 180),
        ["Teal"] = FromRgb(0, 128, 128),
        ["Wheat"] = FromRgb(245, 222, 179),
        ["White"] = FromRgb(255, 255, 255),
        ["WhiteSmoke"] = FromRgb(245, 245, 245),
        ["Yellow"] = FromRgb(255, 255, 0),
        ["YellowGreen"] = FromRgb(154, 205, 50)
    };

    /// <summary>Creates an RGB or RGBA Markdown color.</summary>
    public MarkdownColor(byte r, byte g, byte b, byte a = 255) {
        R = r;
        G = g;
        B = b;
        A = a;
    }

    /// <summary>Red component.</summary>
    public byte R { get; }

    /// <summary>Green component.</summary>
    public byte G { get; }

    /// <summary>Blue component.</summary>
    public byte B { get; }

    /// <summary>Alpha component.</summary>
    public byte A { get; }

    /// <summary>Creates an opaque color from RGB components.</summary>
    public static MarkdownColor FromRgb(byte r, byte g, byte b) => new MarkdownColor(r, g, b);

    /// <summary>Parses a named color, #RGB, RGB, #RRGGBB, RRGGBB, #RRGGBBAA, or RRGGBBAA value.</summary>
    public static MarkdownColor Parse(string value) {
        if (TryParse(value, out MarkdownColor color)) {
            return color;
        }

        throw new FormatException($"Invalid Markdown color value: '{value}'.");
    }

    /// <summary>Tries to parse a named or hexadecimal color.</summary>
    public static bool TryParse(string? value, out MarkdownColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string trimmed = value!.Trim();
        if (NamedColors.TryGetValue(trimmed, out color)) {
            return true;
        }

        return TryParseHex(trimmed, out color);
    }

    /// <summary>Returns the color as #RRGGBB or #RRGGBBAA.</summary>
    public string ToCssColor() => A == 255 ? "#" + ToRgbHex() : "#" + ToRgbaHex();

    /// <summary>Returns RRGGBB.</summary>
    public string ToRgbHex() => $"{R:x2}{G:x2}{B:x2}";

    /// <summary>Returns RRGGBBAA.</summary>
    public string ToRgbaHex() => $"{R:x2}{G:x2}{B:x2}{A:x2}";

    /// <inheritdoc />
    public override string ToString() => ToCssColor();

    /// <inheritdoc />
    public bool Equals(MarkdownColor other) => R == other.R && G == other.G && B == other.B && A == other.A;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is MarkdownColor other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (R << 24) | (G << 16) | (B << 8) | A;

    /// <summary>Equality operator.</summary>
    public static bool operator ==(MarkdownColor left, MarkdownColor right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(MarkdownColor left, MarkdownColor right) => !left.Equals(right);

    private static bool TryParseHex(string value, out MarkdownColor color) {
        color = default;
        string hex = value.Trim();
        if (hex.StartsWith("#", StringComparison.Ordinal)) {
            hex = hex.Substring(1);
        }

        if (hex.Length == 3) {
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
        } else if (hex.Length == 4) {
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2], hex[3], hex[3] });
        }

        if (hex.Length != 6 && hex.Length != 8) {
            return false;
        }

        if (!uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint packed)) {
            return false;
        }

        color = hex.Length == 6
            ? new MarkdownColor((byte)((packed >> 16) & 0xFF), (byte)((packed >> 8) & 0xFF), (byte)(packed & 0xFF))
            : new MarkdownColor((byte)((packed >> 24) & 0xFF), (byte)((packed >> 16) & 0xFF), (byte)((packed >> 8) & 0xFF), (byte)(packed & 0xFF));

        return true;
    }
}
