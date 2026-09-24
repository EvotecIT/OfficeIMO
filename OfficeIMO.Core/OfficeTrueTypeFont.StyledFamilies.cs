using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private static readonly Dictionary<string, StyledFontFamilyResolution> StyledFontFamilyCache = new(StringComparer.OrdinalIgnoreCase);

    // Windows abbreviates the file names of a family's faces, so they cannot be found by family name.
    // Each entry lists the regular, bold, italic and bold-italic files.
    private static readonly Dictionary<string, string?[]> WindowsStyledFontFiles = new(StringComparer.Ordinal) {
        ["arial"] = new[] { "arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf" },
        ["timesnewroman"] = new[] { "times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf" },
        ["couriernew"] = new[] { "cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf" },
        ["calibri"] = new[] { "calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf" },
        ["cambria"] = new[] { "cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf" },
        ["segoeui"] = new[] { "segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf" },
        ["consolas"] = new[] { "consola.ttf", "consolab.ttf", "consolai.ttf", "consolaz.ttf" },
        ["tahoma"] = new[] { "tahoma.ttf", "tahomabd.ttf", null, null },
        ["verdana"] = new[] { "verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf" },
        ["georgia"] = new[] { "georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf" },
        ["trebuchetms"] = new[] { "trebuc.ttf", "trebucbd.ttf", "trebucit.ttf", "trebucbi.ttf" }
    };

    /// <summary>
    /// Loads the installed face of the first resolvable family in a font-family list that carries the requested
    /// bold and italic style, such as Arial Bold for a bold Arial request.
    /// </summary>
    /// <param name="fontFamily">CSS/Office-style font-family fallback list.</param>
    /// <param name="style">Requested style; only bold and italic select a face.</param>
    /// <param name="resolvedStyle">The requested bold and italic bits the returned face provides; the caller simulates the rest.</param>
    /// <returns>The styled face, the family's regular face when it has no styled face, or <see langword="null"/>.</returns>
    internal static OfficeTrueTypeFont? TryLoadFontFamily(string? fontFamily, OfficeFontStyle style, out OfficeFontStyle resolvedStyle) {
        OfficeFontStyle requested = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        resolvedStyle = OfficeFontStyle.Regular;
        if (requested == OfficeFontStyle.Regular || string.IsNullOrEmpty(fontFamily)) {
            return TryLoadFontFamily(fontFamily);
        }

        string cacheKey = NormalizeFontFamilyCacheKey(fontFamily) + "#" + ((int)requested).ToString(System.Globalization.CultureInfo.InvariantCulture);
        StyledFontFamilyResolution resolved;
        bool cached;
        lock (FontCacheLock) {
            cached = StyledFontFamilyCache.TryGetValue(cacheKey, out resolved);
        }

        if (!cached) {
            resolved = ResolveStyledFontFamily(fontFamily!, requested);
            lock (FontCacheLock) {
                if (StyledFontFamilyCache.Count >= MaxFontCacheEntries) StyledFontFamilyCache.Clear();
                StyledFontFamilyCache[cacheKey] = resolved;
            }
        }

        if (resolved.Font != null) {
            resolvedStyle = resolved.Style;
            return resolved.Font;
        }

        return TryLoadFontFamily(fontFamily);
    }

    /// <summary>Resolves an installed face that covers the actual text before accepting a styled or fallback face.</summary>
    internal static OfficeTrueTypeFont? TryLoadFontFamilyForText(string? fontFamily, OfficeFontStyle style,
        string? text, out OfficeFontStyle resolvedStyle) {
        resolvedStyle = OfficeFontStyle.Regular;
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(fontFamily))
            return TryLoadFontFamily(fontFamily, style, out resolvedStyle);

        foreach (string family in ExpandFontFamilyFallbacks(fontFamily!)) {
            OfficeTrueTypeFont? styled = TryLoadFontFamily(family, style, out OfficeFontStyle faceStyle);
            OfficeFontStyle requested = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
            bool styledCoversText = styled != null && styled.HasGlyphs(text!);
            if (styledCoversText && faceStyle == requested) {
                resolvedStyle = faceStyle;
                return styled;
            }
            if (requested != OfficeFontStyle.Regular) {
                string key = NormalizeFontFamilyKey(family);
                OfficeTrueTypeFont? partial = null;
                foreach (string path in CandidateStyledFamilyPaths(key, requested)) {
                    foreach (OfficeTrueTypeFont face in LoadFaces(path)) {
                        OfficeFontStyle candidateStyle = face.FaceStyle;
                        if (candidateStyle == OfficeFontStyle.Regular || (candidateStyle & ~requested) != 0 ||
                            !face.HasFamilyKey(key) || !face.HasGlyphs(text!)) continue;
                        if (candidateStyle == requested) {
                            resolvedStyle = candidateStyle;
                            return face;
                        }
                        partial ??= face;
                    }
                }
                if (partial != null) {
                    resolvedStyle = partial.FaceStyle;
                    return partial;
                }
            }
            if (styledCoversText && faceStyle != OfficeFontStyle.Regular) {
                resolvedStyle = faceStyle;
                return styled;
            }
            OfficeTrueTypeFont? regular = TryLoadFontFamily(family);
            if (regular != null && regular.HasGlyphs(text!)) return regular;
            string regularKey = NormalizeFontFamilyKey(family);
            foreach (string path in CandidateFamilyPaths(family)) {
                foreach (OfficeTrueTypeFont face in LoadFaces(path)) {
                    if (face.FaceStyle == OfficeFontStyle.Regular && face.HasFamilyKey(regularKey) && face.HasGlyphs(text!))
                        return face;
                }
            }
        }
        return null;
    }

    /// <summary>Gets the bold and italic style the face declares in its <c>head</c> table.</summary>
    internal OfficeFontStyle FaceStyle {
        get {
            if (!InBounds(_head + 44, 2)) return OfficeFontStyle.Regular;
            ushort macStyle = ReadUInt16(_data, _head + 44);
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if ((macStyle & 1) != 0) style |= OfficeFontStyle.Bold;
            if ((macStyle & 2) != 0) style |= OfficeFontStyle.Italic;
            return style;
        }
    }

    private static StyledFontFamilyResolution ResolveStyledFontFamily(string fontFamily, OfficeFontStyle requested) {
        foreach (string family in ExpandFontFamilyFallbacks(fontFamily)) {
            string key = NormalizeFontFamilyKey(family);
            OfficeTrueTypeFont? partial = null;
            foreach (string path in CandidateStyledFamilyPaths(key, requested)) {
                foreach (OfficeTrueTypeFont face in LoadFaces(path)) {
                    OfficeFontStyle faceStyle = face.FaceStyle;
                    // A face must not add a style that was not requested, such as bold italic for bold.
                    if (faceStyle == OfficeFontStyle.Regular || (faceStyle & ~requested) != 0) continue;
                    if (!face.HasFamilyKey(key) || !face.HasGlyphs("OfficeIMO 0123456789")) continue;
                    if (faceStyle == requested) return new StyledFontFamilyResolution(face, faceStyle);
                    partial ??= face;
                }
            }

            if (partial != null) return new StyledFontFamilyResolution(partial, partial.FaceStyle);
            // The first installed family wins; its regular face is used with a simulated style.
            if (HasInstalledFace(family)) break;
        }

        return new StyledFontFamilyResolution(null, OfficeFontStyle.Regular);
    }

    private static bool HasInstalledFace(string family) {
        foreach (string path in CandidateFamilyPaths(family)) {
            OfficeTrueTypeFont? font = TryLoad(path, null, family) ?? TryLoad(path);
            if (font != null && font.HasGlyphs("OfficeIMO 0123456789")) return true;
        }

        return false;
    }

    private static IEnumerable<string> CandidateStyledFamilyPaths(string key, OfficeFontStyle requested) {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(windows) && WindowsStyledFontFiles.TryGetValue(key, out string?[]? files)) {
            string fonts = Path.Combine(windows, "Fonts");
            string? exact = files[(int)requested];
            if (exact != null) yield return Path.Combine(fonts, exact);
            if (requested == (OfficeFontStyle.Bold | OfficeFontStyle.Italic)) {
                if (files[(int)OfficeFontStyle.Bold] is string bold) yield return Path.Combine(fonts, bold);
                if (files[(int)OfficeFontStyle.Italic] is string italic) yield return Path.Combine(fonts, italic);
            }
        }

        // Other platforms name faces after the family ("Arial Bold.ttf", "LiberationSans-Bold.ttf") or
        // keep them in one collection ("Helvetica.ttc").
        foreach (string path in CandidateFontDirectoryPaths(key)) {
            yield return path;
        }
    }

    private static IEnumerable<OfficeTrueTypeFont> LoadFaces(string path) {
        int count = ReadCollectionFaceCount(path);
        if (count == 0) {
            OfficeTrueTypeFont? font = TryLoad(path);
            if (font != null) yield return font;
            yield break;
        }

        for (int index = 0; index < count; index++) {
            OfficeTrueTypeFont? face = TryLoad(path, index);
            if (face != null) yield return face;
        }
    }

    // Returns the number of faces in a font collection, or zero for a standalone font or unreadable file.
    private static int ReadCollectionFaceCount(string path) {
        try {
            if (!File.Exists(path)) return 0;
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var header = new byte[12];
            int read = 0;
            while (read < header.Length) {
                int chunk = stream.Read(header, read, header.Length - read);
                if (chunk == 0) return 0;
                read += chunk;
            }

            if (ReadUInt32(header, 0) != 0x74746366) return 0;
            uint count = ReadUInt32(header, 8);
            return count == 0 || count > MaxTrueTypeCollectionFonts ? 0 : (int)count;
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        } catch (ArgumentException) {
        } catch (NotSupportedException) {
        }

        return 0;
    }

    // Compares the face's family and typographic family names with a normalized family key, so Arial Narrow
    // or Arial Black is not taken for Arial.
    private bool HasFamilyKey(string key) {
        if (_name < 0 || _name + 6 > _data.Length) return false;
        var count = ReadUInt16(_data, _name + 2);
        var stringOffset = _name + ReadUInt16(_data, _name + 4);
        for (var i = 0; i < count; i++) {
            var record = _name + 6 + i * 12;
            if (record + 12 > _data.Length) return false;
            var nameId = ReadUInt16(_data, record + 6);
            if (nameId != 1 && nameId != 16) continue;
            var platform = ReadUInt16(_data, record);
            var length = ReadUInt16(_data, record + 8);
            var offset = stringOffset + ReadUInt16(_data, record + 10);
            if (offset < 0 || length == 0 || offset + length > _data.Length) continue;
            if (NormalizeFontFamilyKey(DecodeName(platform, offset, length)) == key) return true;
        }

        return false;
    }

    private readonly struct StyledFontFamilyResolution {
        public StyledFontFamilyResolution(OfficeTrueTypeFont? font, OfficeFontStyle style) {
            Font = font;
            Style = style;
        }

        public OfficeTrueTypeFont? Font { get; }

        public OfficeFontStyle Style { get; }
    }
}
