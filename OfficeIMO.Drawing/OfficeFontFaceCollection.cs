using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Scoped caller-supplied TrueType faces used by drawing measurement, rasterization, and SVG export.
/// Direct sfnt and WOFF 1 containers are accepted and normalized to OpenType bytes.
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
            throw new ArgumentException("The supplied bytes are not a supported TrueType outline font container.", nameof(data));
        }

        return this;
    }

    /// <summary>Attempts to add or replace one family/style face without throwing for unsupported font data.</summary>
    public bool TryAdd(string? familyName, byte[]? data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        return TryAdd(familyName, data, style, OfficeFontUnicodeRangeSet.All);
    }

    /// <summary>
    /// Adds or replaces one unicode-range-constrained family/style face.
    /// A deterministic internal resource family is assigned when the range does not cover all Unicode scalars.
    /// </summary>
    public OfficeFontFaceCollection Add(
        string familyName,
        byte[] data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet unicodeRanges) {
        if (!TryAdd(familyName, data, style, unicodeRanges)) {
            throw new ArgumentException("The supplied bytes are not a supported TrueType outline font container.", nameof(data));
        }

        return this;
    }

    /// <summary>
    /// Attempts to add or replace one unicode-range-constrained family/style face.
    /// A deterministic internal resource family is assigned when the range does not cover all Unicode scalars.
    /// </summary>
    public bool TryAdd(
        string? familyName,
        byte[]? data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet? unicodeRanges) {
        OfficeFontUnicodeRangeSet normalizedRanges = unicodeRanges ?? OfficeFontUnicodeRangeSet.All;
        string? resourceFamilyName = normalizedRanges.IsAll || string.IsNullOrWhiteSpace(familyName)
            ? familyName
            : CreateResourceFamilyName(familyName!.Trim(), style, normalizedRanges);
        return TryAddCore(familyName, data, style, normalizedRanges, resourceFamilyName);
    }

    private bool TryAddCore(
        string? familyName,
        byte[]? data,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet unicodeRanges,
        string? resourceFamilyName) {
        if (string.IsNullOrWhiteSpace(familyName) || data == null || data.Length == 0) {
            return false;
        }

        if (!OfficeFontContainerDecoder.TryDecodeToOpenType(
                data,
                out byte[] openTypeData,
                out _,
                out _)) {
            return false;
        }
        OfficeTrueTypeFont? parsed = OfficeTrueTypeFont.TryLoad(openTypeData);
        if (parsed == null) {
            return false;
        }

        string normalizedFamily = familyName!.Trim();
        string normalizedResourceFamily = string.IsNullOrWhiteSpace(resourceFamilyName)
            ? normalizedFamily
            : resourceFamilyName!.Trim();
        OfficeFontUnicodeRangeSet normalizedRanges = unicodeRanges;
        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        for (int index = _faces.Count - 1; index >= 0; index--) {
            OfficeFontFace existing = _faces[index];
            if (existing.Style == normalizedStyle
                && string.Equals(existing.FamilyName, normalizedFamily, StringComparison.OrdinalIgnoreCase)
                && string.Equals(existing.ResourceFamilyName, normalizedResourceFamily, StringComparison.OrdinalIgnoreCase)) {
                _faces[index] = new OfficeFontFace(
                    normalizedFamily,
                    normalizedResourceFamily,
                    openTypeData,
                    normalizedStyle,
                    normalizedRanges,
                    parsed);
                return true;
            }
        }

        _faces.Add(new OfficeFontFace(
            normalizedFamily,
            normalizedResourceFamily,
            openTypeData,
            normalizedStyle,
            normalizedRanges,
            parsed));
        return true;
    }

    /// <summary>Adds independent copies of all faces from another collection.</summary>
    public OfficeFontFaceCollection AddRange(OfficeFontFaceCollection? fonts) {
        if (fonts == null || ReferenceEquals(fonts, this)) {
            return this;
        }

        foreach (OfficeFontFace face in fonts.Faces) {
            TryAddCore(face.FamilyName, face.DataSnapshot, face.Style, face.UnicodeRanges, face.ResourceFamilyName);
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

    internal bool TryMeasureTextElements(
        string text,
        IReadOnlyList<string> elements,
        double fontSize,
        string? familyNames,
        OfficeFontStyle style,
        out IReadOnlyList<double> widths) {
        widths = Array.Empty<double>();
        if (string.IsNullOrEmpty(text) || elements.Count == 0 || fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
            return false;
        }

        OfficeTrueTypeFont? font = ResolveForText(text, familyNames, style, out OfficeFontStyle _);
        if (font == null) return false;

        widths = font.MeasureTextElements(elements, fontSize);
        return true;
    }

    /// <summary>
    /// Splits text into grapheme-safe runs using the first scoped family whose selected face covers each text element.
    /// Unresolved elements retain the original family list for platform or adapter fallback.
    /// </summary>
    public IReadOnlyList<OfficeFontFallbackRun> PlanFallbackRuns(string? text, string? familyNames, OfficeFontStyle style = OfficeFontStyle.Regular) {
        if (string.IsNullOrEmpty(text)) return Array.Empty<OfficeFontFallbackRun>();

        string requestedFamilies = familyNames?.Trim() ?? string.Empty;
        IReadOnlyList<OfficeFontFace> candidates = ResolveFallbackCandidates(requestedFamilies, style);
        var resolvedFamilies = new Dictionary<string, string>(StringComparer.Ordinal);
        var runs = new List<OfficeFontFallbackRun>();
        var currentText = new StringBuilder();
        string? currentFamily = null;
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            string family;
            if (IsWhitespace(element) && currentFamily != null) {
                family = currentFamily;
            } else {
                if (!resolvedFamilies.TryGetValue(element, out string? resolvedFamily)) {
                    OfficeFontFace? face = null;
                    for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
                        if (!candidates[candidateIndex].Covers(element)) continue;
                        face = candidates[candidateIndex];
                        break;
                    }
                    resolvedFamily = face?.ResourceFamilyName ?? requestedFamilies;
                    resolvedFamilies.Add(element, resolvedFamily);
                }
                family = resolvedFamily;
            }
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

    private IReadOnlyList<OfficeFontFace> ResolveFallbackCandidates(string familyNames, OfficeFontStyle style) {
        if (string.IsNullOrWhiteSpace(familyNames) || _faces.Count == 0) return Array.Empty<OfficeFontFace>();

        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        var result = new List<OfficeFontFace>();
        var added = new HashSet<OfficeFontFace>();
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            var exact = new List<OfficeFontFace>();
            var regular = new List<OfficeFontFace>();
            var available = new List<OfficeFontFace>();
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!MatchesFamily(face, family)) continue;
                available.Add(face);
                if (face.Style == normalizedStyle) exact.Add(face);
                if (face.Style == OfficeFontStyle.Regular) regular.Add(face);
            }
            IReadOnlyList<OfficeFontFace> preferred = exact.Count > 0
                ? exact
                : regular.Count > 0
                    ? regular
                    : available;
            foreach (OfficeFontFace face in preferred) {
                if (added.Add(face)) result.Add(face);
            }
        }
        return result;
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
                if (!MatchesFamily(face, family)) {
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
                if (!MatchesFamily(face, family) || !face.Covers(text)) continue;
                first ??= face;
                if (face.Style == normalizedStyle) exact ??= face;
                if (face.Style == OfficeFontStyle.Regular) regular ??= face;
            }

            OfficeFontFace? preferred = exact ?? regular ?? first;
            if (preferred == null) continue;
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

    private static bool MatchesFamily(OfficeFontFace face, string family) =>
        string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)
        || string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase);

    private static string CreateResourceFamilyName(
        string familyName,
        OfficeFontStyle style,
        OfficeFontUnicodeRangeSet ranges) {
        string value = ((int)OfficeFontFace.NormalizeStyle(style)).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "|"
            + ranges.ToStableKey();
        uint hash = 2166136261;
        for (int index = 0; index < value.Length; index++) {
            hash ^= value[index];
            hash *= 16777619;
        }
        return familyName + "__officeimo_" + hash.ToString("x8", System.Globalization.CultureInfo.InvariantCulture);
    }
}
