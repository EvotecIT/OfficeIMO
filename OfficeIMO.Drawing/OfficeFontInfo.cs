using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Small immutable font descriptor used by OfficeIMO packages without taking a dependency on a font engine.
/// </summary>
public readonly struct OfficeFontInfo : IEquatable<OfficeFontInfo> {
    /// <summary>
    /// Creates a font descriptor.
    /// </summary>
    public OfficeFontInfo(string? familyName, double size = 11.0, OfficeFontStyle style = OfficeFontStyle.Regular) {
        FamilyName = familyName ?? string.Empty;
        Size = size;
        Style = style;
    }

    /// <summary>Font family name, when known.</summary>
    public string FamilyName { get; }

    /// <summary>Font size in points.</summary>
    public double Size { get; }

    /// <summary>Font style flags.</summary>
    public OfficeFontStyle Style { get; }

    /// <summary>Whether the descriptor includes bold styling.</summary>
    public bool IsBold => (Style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold;

    /// <summary>Whether the descriptor includes italic styling.</summary>
    public bool IsItalic => (Style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic;

    /// <summary>Default Office font descriptor.</summary>
    public static OfficeFontInfo Default => new OfficeFontInfo("Calibri", 11.0);

    /// <summary>Creates a copy with a different family name.</summary>
    public OfficeFontInfo WithFamilyName(string? familyName) => new OfficeFontInfo(familyName, Size, Style);

    /// <summary>Creates a copy with a different point size.</summary>
    public OfficeFontInfo WithSize(double size) => new OfficeFontInfo(FamilyName, size, Style);

    /// <summary>Creates a copy with different style flags.</summary>
    public OfficeFontInfo WithStyle(OfficeFontStyle style) => new OfficeFontInfo(FamilyName, Size, style);

    /// <inheritdoc />
    public bool Equals(OfficeFontInfo other) =>
        string.Equals(FamilyName, other.FamilyName, StringComparison.Ordinal) &&
        Size.Equals(other.Size) &&
        Style == other.Style;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeFontInfo other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(FamilyName);
            hash = (hash * 31) + Size.GetHashCode();
            hash = (hash * 31) + Style.GetHashCode();
            return hash;
        }
    }

    /// <inheritdoc />
    public override string ToString() {
        var name = string.IsNullOrWhiteSpace(FamilyName) ? "(unspecified)" : FamilyName;
        var style = Style == OfficeFontStyle.Regular ? "Regular" : Style.ToString();
        return string.Format(CultureInfo.InvariantCulture, "{0}, {1:0.##}pt, {2}", name, Size, style);
    }

    /// <summary>Equality operator.</summary>
    public static bool operator ==(OfficeFontInfo left, OfficeFontInfo right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(OfficeFontInfo left, OfficeFontInfo right) => !left.Equals(right);
}
