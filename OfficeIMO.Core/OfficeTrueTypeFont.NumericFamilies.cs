using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private static readonly Dictionary<string, IReadOnlyList<FontFamilyResolution>> NumericFamilyCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly OfficeFontFaceDescriptor _faceDescriptor;

    /// <summary>Face attributes declared by the font's OS/2 and head tables.</summary>
    public OfficeFontFaceDescriptor FaceDescriptor => _faceDescriptor;

    /// <summary>Resolves an installed family using the same numeric face ordering as document-scoped fonts.</summary>
    public static OfficeTrueTypeFont? TryLoadFontFamilyForText(string? families, OfficeFontFaceDescriptor descriptor,
        string? text, out OfficeFontStyle resolvedStyle) {
        resolvedStyle = OfficeFontStyle.Regular;
        if (string.IsNullOrWhiteSpace(families)) return TryLoadDefault();
        foreach (string family in ExpandFontFamilyFallbacks(families!)) {
            foreach (FontFamilyResolution resolved in OrderedNumericFamily(family, descriptor)) {
                OfficeTrueTypeFont candidate = resolved.Font!;
                if (!string.IsNullOrEmpty(text) && !candidate.HasGlyphs(text!)) continue;
                resolvedStyle = candidate.FaceDescriptor.ToStyle();
                return candidate;
            }
        }
        return null;
    }

    private static List<FontFamilyResolution> OrderedNumericFamily(string family, OfficeFontFaceDescriptor descriptor) {
        var candidates = new List<FontFamilyResolution>(GetNumericFamily(family));
        candidates.Sort((left, right) => OfficeFontFaceMatcher.Compare(left.Font!.FaceDescriptor, right.Font!.FaceDescriptor, descriptor));
        return candidates;
    }

    private static IReadOnlyList<FontFamilyResolution> GetNumericFamily(string family) {
        string key = NormalizeFontFamilyKey(family);
        lock (FontCacheLock) { if (NumericFamilyCache.TryGetValue(key, out var cached)) return cached; }
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<FontFamilyResolution>();
        foreach (OfficeFontStyle style in new[] { OfficeFontStyle.Regular, OfficeFontStyle.Bold, OfficeFontStyle.Italic, OfficeFontStyle.Bold | OfficeFontStyle.Italic }) {
            foreach (string path in CandidateStyledFamilyPaths(key, style)) paths.Add(path);
        }
        // Abbreviated Windows filenames include faces beyond the four legacy style flags.
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (key == "segoeui" && windows.Length > 0) {
            foreach (string file in new[] { "segoeuil.ttf", "seguili.ttf", "seguisl.ttf", "seguisli.ttf", "seguisb.ttf", "seguisbi.ttf", "seguibl.ttf", "seguibli.ttf" })
                paths.Add(Path.Combine(windows, "Fonts", file));
        }
        foreach (string path in paths) foreach (OfficeTrueTypeFont font in LoadFaces(path)) {
            if (font.HasFamilyKey(key)) result.Add(new FontFamilyResolution(font, path));
        }
        lock (FontCacheLock) {
            if (NumericFamilyCache.Count >= 32) NumericFamilyCache.Clear();
            NumericFamilyCache[key] = result.AsReadOnly();
        }
        return result;
    }

    private static OfficeFontFaceDescriptor ReadFaceDescriptor(byte[] data, Dictionary<string, int> tables, IReadOnlyDictionary<string, int> lengths, OfficeFontStyle style) {
        int weight = (style & OfficeFontStyle.Bold) != 0 ? 700 : 400;
        double stretch = 100D;
        OfficeFontSlant slant = (style & OfficeFontStyle.Italic) != 0 ? OfficeFontSlant.Italic : OfficeFontSlant.Normal;
        if (tables.TryGetValue("OS/2", out int offset) && lengths.TryGetValue("OS/2", out int length) && length >= 8) {
            int declaredWeight = ReadUInt16(data, offset + 4), width = ReadUInt16(data, offset + 6);
            if (declaredWeight >= 1 && declaredWeight <= 1000) weight = declaredWeight;
            if (width >= 1 && width <= 9) stretch = new[] { 50D, 62.5D, 75D, 87.5D, 100D, 112.5D, 125D, 150D, 200D }[width - 1];
            if (length >= 64) {
                int selection = ReadUInt16(data, offset + 62);
                if ((selection & 512) != 0) slant = OfficeFontSlant.Oblique;
                else if ((selection & 1) != 0) slant = OfficeFontSlant.Italic;
            }
        }
        return new OfficeFontFaceDescriptor(weight, stretch, slant);
    }
}
