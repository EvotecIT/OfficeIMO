using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Small immutable RGBA color value used by OfficeIMO packages without external image dependencies.
/// </summary>
public readonly partial struct OfficeColor : IEquatable<OfficeColor> {
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

    /// <summary>Returns a CSS hexadecimal color using #RRGGBB or #RRGGBBAA notation.</summary>
    public string ToCssColor() => A == 255
        ? "#" + ToRgbHex().ToLowerInvariant()
        : "#" + ToHex().ToLowerInvariant();

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

    /// <summary>Gets fully transparent black (<c>#00000000</c>).</summary>
    public static OfficeColor Transparent => FromRgba(0, 0, 0, 0);
    /// <summary>Gets the AliceBlue named color (<c>#F0F8FF</c>).</summary>
    public static OfficeColor AliceBlue => FromRgb(240, 248, 255);
    /// <summary>Gets the Aqua named color (<c>#00FFFF</c>).</summary>
    public static OfficeColor Aqua => FromRgb(0, 255, 255);
    /// <summary>Gets the Aquamarine named color (<c>#7FFFD4</c>).</summary>
    public static OfficeColor Aquamarine => FromRgb(127, 255, 212);
    /// <summary>Gets the Azure named color (<c>#F0FFFF</c>).</summary>
    public static OfficeColor Azure => FromRgb(240, 255, 255);
    /// <summary>Gets the Beige named color (<c>#F5F5DC</c>).</summary>
    public static OfficeColor Beige => FromRgb(245, 245, 220);
    /// <summary>Gets the Black named color (<c>#000000</c>).</summary>
    public static OfficeColor Black => FromRgb(0, 0, 0);
    /// <summary>Gets the Blue named color (<c>#0000FF</c>).</summary>
    public static OfficeColor Blue => FromRgb(0, 0, 255);
    /// <summary>Gets the BlueViolet named color (<c>#8A2BE2</c>).</summary>
    public static OfficeColor BlueViolet => FromRgb(138, 43, 226);
    /// <summary>Gets the Brown named color (<c>#A52A2A</c>).</summary>
    public static OfficeColor Brown => FromRgb(165, 42, 42);
    /// <summary>Gets the Coral named color (<c>#FF7F50</c>).</summary>
    public static OfficeColor Coral => FromRgb(255, 127, 80);
    /// <summary>Gets the CornflowerBlue named color (<c>#6495ED</c>).</summary>
    public static OfficeColor CornflowerBlue => FromRgb(100, 149, 237);
    /// <summary>Gets the Crimson named color (<c>#DC143C</c>).</summary>
    public static OfficeColor Crimson => FromRgb(220, 20, 60);
    /// <summary>Gets the Cyan named color (<c>#00FFFF</c>).</summary>
    public static OfficeColor Cyan => FromRgb(0, 255, 255);
    /// <summary>Gets the DarkBlue named color (<c>#00008B</c>).</summary>
    public static OfficeColor DarkBlue => FromRgb(0, 0, 139);
    /// <summary>Gets the DarkCyan named color (<c>#008B8B</c>).</summary>
    public static OfficeColor DarkCyan => FromRgb(0, 139, 139);
    /// <summary>Gets the DarkGoldenrod named color (<c>#B8860B</c>).</summary>
    public static OfficeColor DarkGoldenrod => FromRgb(184, 134, 11);
    /// <summary>Gets the DarkGray named color (<c>#A9A9A9</c>).</summary>
    public static OfficeColor DarkGray => FromRgb(169, 169, 169);
    /// <summary>Gets the DarkGreen named color (<c>#006400</c>).</summary>
    public static OfficeColor DarkGreen => FromRgb(0, 100, 0);
    /// <summary>Gets the DarkKhaki named color (<c>#BDB76B</c>).</summary>
    public static OfficeColor DarkKhaki => FromRgb(189, 183, 107);
    /// <summary>Gets the DarkMagenta named color (<c>#8B008B</c>).</summary>
    public static OfficeColor DarkMagenta => FromRgb(139, 0, 139);
    /// <summary>Gets the DarkOliveGreen named color (<c>#556B2F</c>).</summary>
    public static OfficeColor DarkOliveGreen => FromRgb(85, 107, 47);
    /// <summary>Gets the DarkOrange named color (<c>#FF8C00</c>).</summary>
    public static OfficeColor DarkOrange => FromRgb(255, 140, 0);
    /// <summary>Gets the DarkRed named color (<c>#8B0000</c>).</summary>
    public static OfficeColor DarkRed => FromRgb(139, 0, 0);
    /// <summary>Gets the DarkSalmon named color (<c>#E9967A</c>).</summary>
    public static OfficeColor DarkSalmon => FromRgb(233, 150, 122);
    /// <summary>Gets the DarkSeaGreen named color (<c>#8FBC8F</c>).</summary>
    public static OfficeColor DarkSeaGreen => FromRgb(143, 188, 143);
    /// <summary>Gets the DarkSlateBlue named color (<c>#483D8B</c>).</summary>
    public static OfficeColor DarkSlateBlue => FromRgb(72, 61, 139);
    /// <summary>Gets the DeepPink named color (<c>#FF1493</c>).</summary>
    public static OfficeColor DeepPink => FromRgb(255, 20, 147);
    /// <summary>Gets the DimGray named color (<c>#696969</c>).</summary>
    public static OfficeColor DimGray => FromRgb(105, 105, 105);
    /// <summary>Gets the British-spelling alias for <see cref="DimGray"/>.</summary>
    public static OfficeColor DimGrey => DimGray;
    /// <summary>Gets the DodgerBlue named color (<c>#1E90FF</c>).</summary>
    public static OfficeColor DodgerBlue => FromRgb(30, 144, 255);
    /// <summary>Gets the ForestGreen named color (<c>#228B22</c>).</summary>
    public static OfficeColor ForestGreen => FromRgb(34, 139, 34);
    /// <summary>Gets the Gold named color (<c>#FFD700</c>).</summary>
    public static OfficeColor Gold => FromRgb(255, 215, 0);
    /// <summary>Gets the Goldenrod named color (<c>#DAA520</c>).</summary>
    public static OfficeColor Goldenrod => FromRgb(218, 165, 32);
    /// <summary>Gets the Gray named color (<c>#808080</c>).</summary>
    public static OfficeColor Gray => FromRgb(128, 128, 128);
    /// <summary>Gets the Green named color (<c>#008000</c>).</summary>
    public static OfficeColor Green => FromRgb(0, 128, 0);
    /// <summary>Gets the GreenYellow named color (<c>#ADFF2F</c>).</summary>
    public static OfficeColor GreenYellow => FromRgb(173, 255, 47);
    /// <summary>Gets the Honeydew named color (<c>#F0FFF0</c>).</summary>
    public static OfficeColor Honeydew => FromRgb(240, 255, 240);
    /// <summary>Gets the HotPink named color (<c>#FF69B4</c>).</summary>
    public static OfficeColor HotPink => FromRgb(255, 105, 180);
    /// <summary>Gets the IndianRed named color (<c>#CD5C5C</c>).</summary>
    public static OfficeColor IndianRed => FromRgb(205, 92, 92);
    /// <summary>Gets the Indigo named color (<c>#4B0082</c>).</summary>
    public static OfficeColor Indigo => FromRgb(75, 0, 130);
    /// <summary>Gets the Khaki named color (<c>#F0E68C</c>).</summary>
    public static OfficeColor Khaki => FromRgb(240, 230, 140);
    /// <summary>Gets the Lavender named color (<c>#E6E6FA</c>).</summary>
    public static OfficeColor Lavender => FromRgb(230, 230, 250);
    /// <summary>Gets the LightBlue named color (<c>#ADD8E6</c>).</summary>
    public static OfficeColor LightBlue => FromRgb(173, 216, 230);
    /// <summary>Gets the LightCoral named color (<c>#F08080</c>).</summary>
    public static OfficeColor LightCoral => FromRgb(240, 128, 128);
    /// <summary>Gets the LightCyan named color (<c>#E0FFFF</c>).</summary>
    public static OfficeColor LightCyan => FromRgb(224, 255, 255);
    /// <summary>Gets the LightGoldenrodYellow named color (<c>#FAFAD2</c>).</summary>
    public static OfficeColor LightGoldenrodYellow => FromRgb(250, 250, 210);
    /// <summary>Gets the LightGray named color (<c>#D3D3D3</c>).</summary>
    public static OfficeColor LightGray => FromRgb(211, 211, 211);
    /// <summary>Gets the LightGreen named color (<c>#90EE90</c>).</summary>
    public static OfficeColor LightGreen => FromRgb(144, 238, 144);
    /// <summary>Gets the LightPink named color (<c>#FFB6C1</c>).</summary>
    public static OfficeColor LightPink => FromRgb(255, 182, 193);
    /// <summary>Gets the LightSkyBlue named color (<c>#87CEFA</c>).</summary>
    public static OfficeColor LightSkyBlue => FromRgb(135, 206, 250);
    /// <summary>Gets the LightSteelBlue named color (<c>#B0C4DE</c>).</summary>
    public static OfficeColor LightSteelBlue => FromRgb(176, 196, 222);
    /// <summary>Gets the LightYellow named color (<c>#FFFFE0</c>).</summary>
    public static OfficeColor LightYellow => FromRgb(255, 255, 224);
    /// <summary>Gets the Lime named color (<c>#00FF00</c>).</summary>
    public static OfficeColor Lime => FromRgb(0, 255, 0);
    /// <summary>Gets the Magenta named color (<c>#FF00FF</c>).</summary>
    public static OfficeColor Magenta => FromRgb(255, 0, 255);
    /// <summary>Gets the Maroon named color (<c>#800000</c>).</summary>
    public static OfficeColor Maroon => FromRgb(128, 0, 0);
    /// <summary>Gets the MediumAquamarine named color (<c>#66CDAA</c>).</summary>
    public static OfficeColor MediumAquamarine => FromRgb(102, 205, 170);
    /// <summary>Gets the MediumPurple named color (<c>#9370DB</c>).</summary>
    public static OfficeColor MediumPurple => FromRgb(147, 112, 219);
    /// <summary>Gets the MediumVioletRed named color (<c>#C71585</c>).</summary>
    public static OfficeColor MediumVioletRed => FromRgb(199, 21, 133);
    /// <summary>Gets the MistyRose named color (<c>#FFE4E1</c>).</summary>
    public static OfficeColor MistyRose => FromRgb(255, 228, 225);
    /// <summary>Gets the Navy named color (<c>#000080</c>).</summary>
    public static OfficeColor Navy => FromRgb(0, 0, 128);
    /// <summary>Gets the Olive named color (<c>#808000</c>).</summary>
    public static OfficeColor Olive => FromRgb(128, 128, 0);
    /// <summary>Gets the Orange named color (<c>#FFA500</c>).</summary>
    public static OfficeColor Orange => FromRgb(255, 165, 0);
    /// <summary>Gets the OrangeRed named color (<c>#FF4500</c>).</summary>
    public static OfficeColor OrangeRed => FromRgb(255, 69, 0);
    /// <summary>Gets the Peru named color (<c>#CD853F</c>).</summary>
    public static OfficeColor Peru => FromRgb(205, 133, 63);
    /// <summary>Gets the Pink named color (<c>#FFC0CB</c>).</summary>
    public static OfficeColor Pink => FromRgb(255, 192, 203);
    /// <summary>Gets the Plum named color (<c>#DDA0DD</c>).</summary>
    public static OfficeColor Plum => FromRgb(221, 160, 221);
    /// <summary>Gets the Purple named color (<c>#800080</c>).</summary>
    public static OfficeColor Purple => FromRgb(128, 0, 128);
    /// <summary>Gets the RebeccaPurple named color (<c>#663399</c>).</summary>
    public static OfficeColor RebeccaPurple => FromRgb(102, 51, 153);
    /// <summary>Gets the Red named color (<c>#FF0000</c>).</summary>
    public static OfficeColor Red => FromRgb(255, 0, 0);
    /// <summary>Gets the SaddleBrown named color (<c>#8B4513</c>).</summary>
    public static OfficeColor SaddleBrown => FromRgb(139, 69, 19);
    /// <summary>Gets the Salmon named color (<c>#FA8072</c>).</summary>
    public static OfficeColor Salmon => FromRgb(250, 128, 114);
    /// <summary>Gets the SandyBrown named color (<c>#F4A460</c>).</summary>
    public static OfficeColor SandyBrown => FromRgb(244, 164, 96);
    /// <summary>Gets the SeaGreen named color (<c>#2E8B57</c>).</summary>
    public static OfficeColor SeaGreen => FromRgb(46, 139, 87);
    /// <summary>Gets the Silver named color (<c>#C0C0C0</c>).</summary>
    public static OfficeColor Silver => FromRgb(192, 192, 192);
    /// <summary>Gets the SkyBlue named color (<c>#87CEEB</c>).</summary>
    public static OfficeColor SkyBlue => FromRgb(135, 206, 235);
    /// <summary>Gets the SteelBlue named color (<c>#4682B4</c>).</summary>
    public static OfficeColor SteelBlue => FromRgb(70, 130, 180);
    /// <summary>Gets the Teal named color (<c>#008080</c>).</summary>
    public static OfficeColor Teal => FromRgb(0, 128, 128);
    /// <summary>Gets the Wheat named color (<c>#F5DEB3</c>).</summary>
    public static OfficeColor Wheat => FromRgb(245, 222, 179);
    /// <summary>Gets the White named color (<c>#FFFFFF</c>).</summary>
    public static OfficeColor White => FromRgb(255, 255, 255);
    /// <summary>Gets the WhiteSmoke named color (<c>#F5F5F5</c>).</summary>
    public static OfficeColor WhiteSmoke => FromRgb(245, 245, 245);
    /// <summary>Gets the Yellow named color (<c>#FFFF00</c>).</summary>
    public static OfficeColor Yellow => FromRgb(255, 255, 0);
    /// <summary>Gets the YellowGreen named color (<c>#9ACD32</c>).</summary>
    public static OfficeColor YellowGreen => FromRgb(154, 205, 50);
}
