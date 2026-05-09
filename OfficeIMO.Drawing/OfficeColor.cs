using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Small immutable RGBA color value used by OfficeIMO packages without external image dependencies.
/// </summary>
public readonly struct OfficeColor : IEquatable<OfficeColor> {
    private static readonly Dictionary<string, OfficeColor> NamedColors = new Dictionary<string, OfficeColor>(StringComparer.OrdinalIgnoreCase) {
        ["Transparent"] = Transparent,
        ["AliceBlue"] = AliceBlue,
        ["AntiqueWhite"] = FromRgb(250, 235, 215),
        ["Aqua"] = Aqua,
        ["Aquamarine"] = Aquamarine,
        ["Azure"] = Azure,
        ["Beige"] = Beige,
        ["Bisque"] = FromRgb(255, 228, 196),
        ["Black"] = Black,
        ["BlanchedAlmond"] = FromRgb(255, 235, 205),
        ["Blue"] = Blue,
        ["BlueViolet"] = BlueViolet,
        ["Brown"] = Brown,
        ["BurlyWood"] = FromRgb(222, 184, 135),
        ["CadetBlue"] = FromRgb(95, 158, 160),
        ["Chartreuse"] = FromRgb(127, 255, 0),
        ["Chocolate"] = FromRgb(210, 105, 30),
        ["Coral"] = Coral,
        ["CornflowerBlue"] = CornflowerBlue,
        ["Cornsilk"] = FromRgb(255, 248, 220),
        ["Crimson"] = Crimson,
        ["Cyan"] = Cyan,
        ["DarkBlue"] = DarkBlue,
        ["DarkCyan"] = DarkCyan,
        ["DarkGoldenrod"] = DarkGoldenrod,
        ["DarkGray"] = DarkGray,
        ["DarkGreen"] = DarkGreen,
        ["DarkGrey"] = DarkGray,
        ["DarkKhaki"] = DarkKhaki,
        ["DarkMagenta"] = DarkMagenta,
        ["DarkOliveGreen"] = DarkOliveGreen,
        ["DarkOrange"] = DarkOrange,
        ["DarkOrchid"] = FromRgb(153, 50, 204),
        ["DarkRed"] = DarkRed,
        ["DarkSalmon"] = DarkSalmon,
        ["DarkSeaGreen"] = DarkSeaGreen,
        ["DarkSlateBlue"] = DarkSlateBlue,
        ["DarkSlateGray"] = FromRgb(47, 79, 79),
        ["DarkSlateGrey"] = FromRgb(47, 79, 79),
        ["DarkTurquoise"] = FromRgb(0, 206, 209),
        ["DarkViolet"] = FromRgb(148, 0, 211),
        ["DeepPink"] = DeepPink,
        ["DeepSkyBlue"] = FromRgb(0, 191, 255),
        ["DimGray"] = DimGray,
        ["DimGrey"] = DimGrey,
        ["DodgerBlue"] = DodgerBlue,
        ["Firebrick"] = FromRgb(178, 34, 34),
        ["FloralWhite"] = FromRgb(255, 250, 240),
        ["ForestGreen"] = ForestGreen,
        ["Fuchsia"] = Magenta,
        ["Gainsboro"] = FromRgb(220, 220, 220),
        ["GhostWhite"] = FromRgb(248, 248, 255),
        ["Gold"] = Gold,
        ["Goldenrod"] = Goldenrod,
        ["Gray"] = Gray,
        ["Grey"] = Gray,
        ["Green"] = Green,
        ["GreenYellow"] = GreenYellow,
        ["Honeydew"] = Honeydew,
        ["HotPink"] = HotPink,
        ["IndianRed"] = IndianRed,
        ["Indigo"] = Indigo,
        ["Ivory"] = FromRgb(255, 255, 240),
        ["Khaki"] = Khaki,
        ["Lavender"] = Lavender,
        ["LavenderBlush"] = FromRgb(255, 240, 245),
        ["LawnGreen"] = FromRgb(124, 252, 0),
        ["LemonChiffon"] = FromRgb(255, 250, 205),
        ["LightBlue"] = LightBlue,
        ["LightCoral"] = LightCoral,
        ["LightCyan"] = LightCyan,
        ["LightGoldenrodYellow"] = LightGoldenrodYellow,
        ["LightGray"] = LightGray,
        ["LightGreen"] = LightGreen,
        ["LightGrey"] = LightGray,
        ["LightPink"] = LightPink,
        ["LightSalmon"] = FromRgb(255, 160, 122),
        ["LightSeaGreen"] = FromRgb(32, 178, 170),
        ["LightSkyBlue"] = LightSkyBlue,
        ["LightSlateGray"] = FromRgb(119, 136, 153),
        ["LightSlateGrey"] = FromRgb(119, 136, 153),
        ["LightSteelBlue"] = LightSteelBlue,
        ["LightYellow"] = LightYellow,
        ["Lime"] = Lime,
        ["LimeGreen"] = FromRgb(50, 205, 50),
        ["Linen"] = FromRgb(250, 240, 230),
        ["Magenta"] = Magenta,
        ["Maroon"] = Maroon,
        ["MediumAquamarine"] = MediumAquamarine,
        ["MediumBlue"] = FromRgb(0, 0, 205),
        ["MediumOrchid"] = FromRgb(186, 85, 211),
        ["MediumPurple"] = MediumPurple,
        ["MediumSeaGreen"] = FromRgb(60, 179, 113),
        ["MediumSlateBlue"] = FromRgb(123, 104, 238),
        ["MediumSpringGreen"] = FromRgb(0, 250, 154),
        ["MediumTurquoise"] = FromRgb(72, 209, 204),
        ["MediumVioletRed"] = MediumVioletRed,
        ["MidnightBlue"] = FromRgb(25, 25, 112),
        ["MintCream"] = FromRgb(245, 255, 250),
        ["MistyRose"] = MistyRose,
        ["Moccasin"] = FromRgb(255, 228, 181),
        ["NavajoWhite"] = FromRgb(255, 222, 173),
        ["Navy"] = Navy,
        ["OldLace"] = FromRgb(253, 245, 230),
        ["Olive"] = Olive,
        ["OliveDrab"] = FromRgb(107, 142, 35),
        ["Orange"] = Orange,
        ["OrangeRed"] = OrangeRed,
        ["Orchid"] = FromRgb(218, 112, 214),
        ["PaleGoldenrod"] = FromRgb(238, 232, 170),
        ["PaleGreen"] = FromRgb(152, 251, 152),
        ["PaleTurquoise"] = FromRgb(175, 238, 238),
        ["PaleVioletRed"] = FromRgb(219, 112, 147),
        ["PapayaWhip"] = FromRgb(255, 239, 213),
        ["PeachPuff"] = FromRgb(255, 218, 185),
        ["Peru"] = Peru,
        ["Pink"] = Pink,
        ["Plum"] = Plum,
        ["PowderBlue"] = FromRgb(176, 224, 230),
        ["Purple"] = Purple,
        ["RebeccaPurple"] = RebeccaPurple,
        ["Red"] = Red,
        ["RosyBrown"] = FromRgb(188, 143, 143),
        ["RoyalBlue"] = FromRgb(65, 105, 225),
        ["SaddleBrown"] = SaddleBrown,
        ["Salmon"] = Salmon,
        ["SandyBrown"] = SandyBrown,
        ["SeaGreen"] = SeaGreen,
        ["SeaShell"] = FromRgb(255, 245, 238),
        ["Sienna"] = FromRgb(160, 82, 45),
        ["Silver"] = Silver,
        ["SkyBlue"] = SkyBlue,
        ["SlateBlue"] = FromRgb(106, 90, 205),
        ["SlateGray"] = FromRgb(112, 128, 144),
        ["SlateGrey"] = FromRgb(112, 128, 144),
        ["Snow"] = FromRgb(255, 250, 250),
        ["SpringGreen"] = FromRgb(0, 255, 127),
        ["SteelBlue"] = SteelBlue,
        ["Tan"] = FromRgb(210, 180, 140),
        ["Teal"] = Teal,
        ["Thistle"] = FromRgb(216, 191, 216),
        ["Tomato"] = FromRgb(255, 99, 71),
        ["Turquoise"] = FromRgb(64, 224, 208),
        ["Violet"] = FromRgb(238, 130, 238),
        ["Wheat"] = Wheat,
        ["White"] = White,
        ["WhiteSmoke"] = WhiteSmoke,
        ["Yellow"] = Yellow,
        ["YellowGreen"] = YellowGreen
    };

    /// <summary>
    /// Creates an RGBA color.
    /// </summary>
    public OfficeColor(byte r, byte g, byte b, byte a = 255) {
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
    public static OfficeColor FromRgb(byte r, byte g, byte b) => new OfficeColor(r, g, b);

    /// <summary>Creates a color from RGBA components.</summary>
    public static OfficeColor FromRgba(byte r, byte g, byte b, byte a) => new OfficeColor(r, g, b, a);

    /// <summary>Parses a named color, #RRGGBB, RRGGBB, #RRGGBBAA, or RRGGBBAA.</summary>
    public static OfficeColor Parse(string value) {
        if (TryParse(value, out var color)) {
            return color;
        }

        throw new FormatException($"Invalid color value: '{value}'.");
    }

    /// <summary>Parses a hexadecimal color value.</summary>
    public static OfficeColor ParseHex(string value) {
        if (TryParseHex(value, out var color)) {
            return color;
        }

        throw new FormatException($"Invalid hex color value: '{value}'.");
    }

    /// <summary>Tries to parse a named or hexadecimal color value.</summary>
    public static bool TryParse(string? value, out OfficeColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var trimmed = value!.Trim();
        if (NamedColors.TryGetValue(trimmed, out color)) {
            return true;
        }

        return TryParseHex(trimmed, out color);
    }

    /// <summary>Tries to parse a hexadecimal color value.</summary>
    public static bool TryParseHex(string? value, out OfficeColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var hex = value!.Trim();
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

        if (!uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var packed)) {
            return false;
        }

        color = hex.Length == 6
            ? FromRgb((byte)((packed >> 16) & 0xFF), (byte)((packed >> 8) & 0xFF), (byte)(packed & 0xFF))
            : FromRgba((byte)((packed >> 24) & 0xFF), (byte)((packed >> 16) & 0xFF), (byte)((packed >> 8) & 0xFF), (byte)(packed & 0xFF));

        return true;
    }

    /// <summary>Returns RRGGBBAA.</summary>
    public string ToHex() => $"{R:X2}{G:X2}{B:X2}{A:X2}";

    /// <summary>Returns RRGGBB.</summary>
    public string ToRgbHex() => $"{R:X2}{G:X2}{B:X2}";

    /// <summary>Returns AARRGGBB.</summary>
    public string ToArgbHex() => $"{A:X2}{R:X2}{G:X2}{B:X2}";

    /// <inheritdoc />
    public override string ToString() => A == 255 ? $"#{ToRgbHex()}" : $"#{ToHex()}";

    /// <inheritdoc />
    public bool Equals(OfficeColor other) => R == other.R && G == other.G && B == other.B && A == other.A;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeColor other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (R << 24) | (G << 16) | (B << 8) | A;

    /// <summary>Equality operator.</summary>
    public static bool operator ==(OfficeColor left, OfficeColor right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(OfficeColor left, OfficeColor right) => !left.Equals(right);

    public static OfficeColor Transparent => FromRgba(0, 0, 0, 0);
    public static OfficeColor AliceBlue => FromRgb(240, 248, 255);
    public static OfficeColor Aqua => FromRgb(0, 255, 255);
    public static OfficeColor Aquamarine => FromRgb(127, 255, 212);
    public static OfficeColor Azure => FromRgb(240, 255, 255);
    public static OfficeColor Beige => FromRgb(245, 245, 220);
    public static OfficeColor Black => FromRgb(0, 0, 0);
    public static OfficeColor Blue => FromRgb(0, 0, 255);
    public static OfficeColor BlueViolet => FromRgb(138, 43, 226);
    public static OfficeColor Brown => FromRgb(165, 42, 42);
    public static OfficeColor Coral => FromRgb(255, 127, 80);
    public static OfficeColor CornflowerBlue => FromRgb(100, 149, 237);
    public static OfficeColor Crimson => FromRgb(220, 20, 60);
    public static OfficeColor Cyan => FromRgb(0, 255, 255);
    public static OfficeColor DarkBlue => FromRgb(0, 0, 139);
    public static OfficeColor DarkCyan => FromRgb(0, 139, 139);
    public static OfficeColor DarkGoldenrod => FromRgb(184, 134, 11);
    public static OfficeColor DarkGray => FromRgb(169, 169, 169);
    public static OfficeColor DarkGreen => FromRgb(0, 100, 0);
    public static OfficeColor DarkKhaki => FromRgb(189, 183, 107);
    public static OfficeColor DarkMagenta => FromRgb(139, 0, 139);
    public static OfficeColor DarkOliveGreen => FromRgb(85, 107, 47);
    public static OfficeColor DarkOrange => FromRgb(255, 140, 0);
    public static OfficeColor DarkRed => FromRgb(139, 0, 0);
    public static OfficeColor DarkSalmon => FromRgb(233, 150, 122);
    public static OfficeColor DarkSeaGreen => FromRgb(143, 188, 143);
    public static OfficeColor DarkSlateBlue => FromRgb(72, 61, 139);
    public static OfficeColor DeepPink => FromRgb(255, 20, 147);
    public static OfficeColor DimGray => FromRgb(105, 105, 105);
    public static OfficeColor DimGrey => DimGray;
    public static OfficeColor DodgerBlue => FromRgb(30, 144, 255);
    public static OfficeColor ForestGreen => FromRgb(34, 139, 34);
    public static OfficeColor Gold => FromRgb(255, 215, 0);
    public static OfficeColor Goldenrod => FromRgb(218, 165, 32);
    public static OfficeColor Gray => FromRgb(128, 128, 128);
    public static OfficeColor Green => FromRgb(0, 128, 0);
    public static OfficeColor GreenYellow => FromRgb(173, 255, 47);
    public static OfficeColor Honeydew => FromRgb(240, 255, 240);
    public static OfficeColor HotPink => FromRgb(255, 105, 180);
    public static OfficeColor IndianRed => FromRgb(205, 92, 92);
    public static OfficeColor Indigo => FromRgb(75, 0, 130);
    public static OfficeColor Khaki => FromRgb(240, 230, 140);
    public static OfficeColor Lavender => FromRgb(230, 230, 250);
    public static OfficeColor LightBlue => FromRgb(173, 216, 230);
    public static OfficeColor LightCoral => FromRgb(240, 128, 128);
    public static OfficeColor LightCyan => FromRgb(224, 255, 255);
    public static OfficeColor LightGoldenrodYellow => FromRgb(250, 250, 210);
    public static OfficeColor LightGray => FromRgb(211, 211, 211);
    public static OfficeColor LightGreen => FromRgb(144, 238, 144);
    public static OfficeColor LightPink => FromRgb(255, 182, 193);
    public static OfficeColor LightSkyBlue => FromRgb(135, 206, 250);
    public static OfficeColor LightSteelBlue => FromRgb(176, 196, 222);
    public static OfficeColor LightYellow => FromRgb(255, 255, 224);
    public static OfficeColor Lime => FromRgb(0, 255, 0);
    public static OfficeColor Magenta => FromRgb(255, 0, 255);
    public static OfficeColor Maroon => FromRgb(128, 0, 0);
    public static OfficeColor MediumAquamarine => FromRgb(102, 205, 170);
    public static OfficeColor MediumPurple => FromRgb(147, 112, 219);
    public static OfficeColor MediumVioletRed => FromRgb(199, 21, 133);
    public static OfficeColor MistyRose => FromRgb(255, 228, 225);
    public static OfficeColor Navy => FromRgb(0, 0, 128);
    public static OfficeColor Olive => FromRgb(128, 128, 0);
    public static OfficeColor Orange => FromRgb(255, 165, 0);
    public static OfficeColor OrangeRed => FromRgb(255, 69, 0);
    public static OfficeColor Peru => FromRgb(205, 133, 63);
    public static OfficeColor Pink => FromRgb(255, 192, 203);
    public static OfficeColor Plum => FromRgb(221, 160, 221);
    public static OfficeColor Purple => FromRgb(128, 0, 128);
    public static OfficeColor RebeccaPurple => FromRgb(102, 51, 153);
    public static OfficeColor Red => FromRgb(255, 0, 0);
    public static OfficeColor SaddleBrown => FromRgb(139, 69, 19);
    public static OfficeColor Salmon => FromRgb(250, 128, 114);
    public static OfficeColor SandyBrown => FromRgb(244, 164, 96);
    public static OfficeColor SeaGreen => FromRgb(46, 139, 87);
    public static OfficeColor Silver => FromRgb(192, 192, 192);
    public static OfficeColor SkyBlue => FromRgb(135, 206, 235);
    public static OfficeColor SteelBlue => FromRgb(70, 130, 180);
    public static OfficeColor Teal => FromRgb(0, 128, 128);
    public static OfficeColor Wheat => FromRgb(245, 222, 179);
    public static OfficeColor White => FromRgb(255, 255, 255);
    public static OfficeColor WhiteSmoke => FromRgb(245, 245, 245);
    public static OfficeColor Yellow => FromRgb(255, 255, 0);
    public static OfficeColor YellowGreen => FromRgb(154, 205, 50);
}
