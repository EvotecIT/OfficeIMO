using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Scoped caller-supplied TrueType faces used by drawing measurement, rasterization, and SVG export.
/// </summary>
public sealed class OfficeFontFaceCollection {
    private readonly List<OfficeFontFace> _faces = new List<OfficeFontFace>();
    private readonly ReadOnlyCollection<OfficeFontFace> _facesView;

    /// <summary>Creates an empty scoped font collection.</summary>
    public OfficeFontFaceCollection() {
        _facesView = new ReadOnlyCollection<OfficeFontFace>(_faces);
    }

    /// <summary>Registered faces in registration order.</summary>
    public IReadOnlyList<OfficeFontFace> Faces => _facesView;

    /// <summary>Adds or replaces one family/style face. Invalid or unsupported font bytes throw.</summary>
    public OfficeFontFaceCollection Add(string familyName, byte[] data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        if (!TryAdd(familyName, data, style)) {
            throw new ArgumentException("The supplied bytes are not a supported TrueType outline font.", nameof(data));
        }

        return this;
    }

    /// <summary>Attempts to add or replace one family/style face without throwing for unsupported font data.</summary>
    public bool TryAdd(string? familyName, byte[]? data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        if (string.IsNullOrWhiteSpace(familyName) || data == null || data.Length == 0) {
            return false;
        }

        OfficeTrueTypeFont? parsed = OfficeTrueTypeFont.TryLoad(data);
        if (parsed == null) {
            return false;
        }

        string normalizedFamily = familyName!.Trim();
        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        for (int index = _faces.Count - 1; index >= 0; index--) {
            OfficeFontFace existing = _faces[index];
            if (existing.Style == normalizedStyle
                && string.Equals(existing.FamilyName, normalizedFamily, StringComparison.OrdinalIgnoreCase)) {
                _faces[index] = new OfficeFontFace(normalizedFamily, data, normalizedStyle, parsed);
                return true;
            }
        }

        _faces.Add(new OfficeFontFace(normalizedFamily, data, normalizedStyle, parsed));
        return true;
    }

    /// <summary>Adds independent copies of all faces from another collection.</summary>
    public OfficeFontFaceCollection AddRange(OfficeFontFaceCollection? fonts) {
        if (fonts == null || ReferenceEquals(fonts, this)) {
            return this;
        }

        foreach (OfficeFontFace face in fonts.Faces) {
            TryAdd(face.FamilyName, face.DataSnapshot, face.Style);
        }

        return this;
    }

    /// <summary>Creates an independent collection snapshot.</summary>
    public OfficeFontFaceCollection Clone() {
        var clone = new OfficeFontFaceCollection();
        foreach (OfficeFontFace face in _faces) {
            clone._faces.Add(face.Clone());
        }

        return clone;
    }

    /// <summary>Attempts to measure text with a matching scoped face.</summary>
    public bool TryMeasureText(string? text, double fontSize, string? familyNames, OfficeFontStyle style, out double width) {
        width = 0D;
        if (string.IsNullOrEmpty(text) || fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
            return false;
        }

        OfficeTrueTypeFont? font = Resolve(familyNames, style);
        if (font == null) {
            return false;
        }

        width = font.Measure(text!, fontSize);
        return true;
    }

    internal OfficeTrueTypeFont? Resolve(string? familyNames, OfficeFontStyle style) {
        if (string.IsNullOrWhiteSpace(familyNames) || _faces.Count == 0) {
            return null;
        }

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        foreach (string rawFamily in familyNames!.Split(',')) {
            string family = CleanFamilyName(rawFamily);
            if (family.Length == 0) {
                continue;
            }

            OfficeFontFace? regular = null;
            OfficeFontFace? first = null;
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                first ??= face;
                if (face.Style == normalizedStyle) {
                    return face.ParsedFont;
                }

                if (face.Style == OfficeFontStyle.Regular) {
                    regular = face;
                }
            }

            if (regular != null) {
                return regular.ParsedFont;
            }

            if (first != null) {
                return first.ParsedFont;
            }
        }

        return null;
    }

    private static string CleanFamilyName(string familyName) {
        string value = familyName.Trim();
        while (value.Length >= 2
               && ((value[0] == '"' && value[value.Length - 1] == '"')
                   || (value[0] == '\'' && value[value.Length - 1] == '\''))) {
            value = value.Substring(1, value.Length - 2).Trim();
        }

        return value;
    }
}
