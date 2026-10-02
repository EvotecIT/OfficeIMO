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
        Face = OfficeFontFaceDescriptor.FromStyle(style);
    }

    /// <summary>Creates a font with exact numeric face attributes and optional text decorations.</summary>
    public OfficeFontInfo(string? familyName, double size, OfficeFontFaceDescriptor face, OfficeFontStyle decorations = OfficeFontStyle.Regular) {
        FamilyName = familyName ?? string.Empty;
        Size = size;
        Face = face;
        Style = face.ToStyle() | (decorations & ~(OfficeFontStyle.Bold | OfficeFontStyle.Italic));
    }

    /// <summary>Font family name, when known.</summary>
    public string FamilyName { get; }

    /// <summary>Font size in points.</summary>
    public double Size { get; }

    /// <summary>Font style flags.</summary>
    public OfficeFontStyle Style { get; }

    /// <summary>Numeric weight, width, and slant used consistently for measurement, painting, and export.</summary>
    public OfficeFontFaceDescriptor Face { get; }

    /// <summary>Whether the descriptor includes bold styling.</summary>
    public bool IsBold => (Style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold;

    /// <summary>Whether the descriptor includes italic styling.</summary>
    public bool IsItalic => (Style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic;

    /// <summary>Whether the descriptor includes underline styling.</summary>
    public bool IsUnderline => (Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline;

    /// <summary>Whether the descriptor includes strikethrough styling.</summary>
    public bool IsStrikethrough => (Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough;

    /// <summary>Default Office font descriptor.</summary>
    public static OfficeFontInfo Default => new OfficeFontInfo("Calibri", 11.0);

    /// <summary>Creates a copy with a different family name.</summary>
    public OfficeFontInfo WithFamilyName(string? familyName) => new OfficeFontInfo(familyName, Size, Face, Style);

    /// <summary>Creates a copy with a different point size.</summary>
    public OfficeFontInfo WithSize(double size) => new OfficeFontInfo(FamilyName, size, Face, Style);

    /// <summary>Creates a copy with different style flags.</summary>
    public OfficeFontInfo WithStyle(OfficeFontStyle style) {
        OfficeFontFaceDescriptor face = Face;
        if ((style & OfficeFontStyle.Bold) != (Style & OfficeFontStyle.Bold)) face = new OfficeFontFaceDescriptor(
            (style & OfficeFontStyle.Bold) != 0 ? 700 : 400, face.StretchPercent, face.Slant, face.ObliqueAngleDegrees);
        if ((style & OfficeFontStyle.Italic) != (Style & OfficeFontStyle.Italic)) face = new OfficeFontFaceDescriptor(
            face.Weight, face.StretchPercent, (style & OfficeFontStyle.Italic) != 0 ? OfficeFontSlant.Italic : OfficeFontSlant.Normal);
        return new OfficeFontInfo(FamilyName, Size, face, style);
    }

    /// <summary>Creates a copy with exact face attributes, retaining text decorations.</summary>
    public OfficeFontInfo WithFace(OfficeFontFaceDescriptor face) => new OfficeFontInfo(FamilyName, Size, face, Style);

    /// <inheritdoc />
    public bool Equals(OfficeFontInfo other) =>
        string.Equals(FamilyName, other.FamilyName, StringComparison.Ordinal) &&
        Size.Equals(other.Size) &&
        Style == other.Style && Face == other.Face;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeFontInfo other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(FamilyName ?? string.Empty);
            hash = (hash * 31) + Size.GetHashCode();
            hash = (hash * 31) + Style.GetHashCode();
            hash = (hash * 31) + Face.GetHashCode();
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
