using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

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

        OfficeTrueTypeFont? font = ResolveForText(text!, familyNames, style, out OfficeFontStyle _);
        if (font == null) {
            return false;
        }

        width = font.Measure(text!, fontSize);
        return true;
    }

    /// <summary>
    /// Splits text into grapheme-safe runs using the first scoped family whose selected face covers each text element.
    /// Unresolved elements retain the original family list for platform or adapter fallback.
    /// </summary>
    public IReadOnlyList<OfficeFontFallbackRun> PlanFallbackRuns(string? text, string? familyNames, OfficeFontStyle style = OfficeFontStyle.Regular) {
        if (string.IsNullOrEmpty(text)) return Array.Empty<OfficeFontFallbackRun>();

        string requestedFamilies = familyNames?.Trim() ?? string.Empty;
        var runs = new List<OfficeFontFallbackRun>();
        var currentText = new StringBuilder();
        string? currentFamily = null;
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            string family = IsWhitespace(element) && currentFamily != null
                ? currentFamily
                : ResolveForText(element, requestedFamilies, style, out OfficeFontFace? face) != null
                    ? face!.FamilyName
                    : requestedFamilies;
            if (currentFamily != null && !string.Equals(currentFamily, family, StringComparison.OrdinalIgnoreCase)) {
                runs.Add(new OfficeFontFallbackRun(currentText.ToString(), currentFamily));
                currentText.Clear();
            }

            currentFamily = family;
            currentText.Append(element);
        }

        if (currentText.Length > 0) runs.Add(new OfficeFontFallbackRun(currentText.ToString(), currentFamily ?? requestedFamilies));
        return runs.AsReadOnly();
    }

    internal OfficeTrueTypeFont? Resolve(string? familyNames, OfficeFontStyle style) {
        return Resolve(familyNames, style, out _);
    }

    internal OfficeTrueTypeFont? Resolve(string? familyNames, OfficeFontStyle style, out OfficeFontStyle resolvedStyle) {
        resolvedStyle = OfficeFontStyle.Regular;
        if (string.IsNullOrEmpty(familyNames) || _faces.Count == 0) {
            return null;
        }

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            OfficeFontFace? regular = null;
            OfficeFontFace? first = null;
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                first ??= face;
                if (face.Style == normalizedStyle) {
                    resolvedStyle = face.Style;
                    return face.ParsedFont;
                }

                if (face.Style == OfficeFontStyle.Regular) {
                    regular = face;
                }
            }

            if (regular != null) {
                resolvedStyle = regular.Style;
                return regular.ParsedFont;
            }

            if (first != null) {
                resolvedStyle = first.Style;
                return first.ParsedFont;
            }
        }

        return null;
    }

    internal OfficeTrueTypeFont? ResolveForText(string text, string? familyNames, OfficeFontStyle style, out OfficeFontStyle resolvedStyle) {
        OfficeTrueTypeFont? font = ResolveForText(text, familyNames, style, out OfficeFontFace? face);
        resolvedStyle = face?.Style ?? OfficeFontStyle.Regular;
        return font;
    }

    private OfficeTrueTypeFont? ResolveForText(string text, string? familyNames, OfficeFontStyle style, out OfficeFontFace? resolvedFace) {
        resolvedFace = null;
        if (string.IsNullOrEmpty(familyNames) || _faces.Count == 0) return null;

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            OfficeFontFace? exact = null;
            OfficeFontFace? regular = null;
            OfficeFontFace? first = null;
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)) continue;
                first ??= face;
                if (face.Style == normalizedStyle) {
                    exact = face;
                    break;
                }
                if (face.Style == OfficeFontStyle.Regular) regular ??= face;
            }

            OfficeFontFace? preferred = exact ?? regular ?? first;
            if (preferred == null || !preferred.ParsedFont.HasGlyphs(text)) continue;
            resolvedFace = preferred;
            return preferred.ParsedFont;
        }

        return null;
    }

    private static bool IsWhitespace(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (!char.IsWhiteSpace(value[index])) return false;
        }
        return value.Length > 0;
    }

}
