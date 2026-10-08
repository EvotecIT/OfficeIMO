using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
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

    /// <summary>Resolves legacy style flags through the canonical numeric installed-face matcher.</summary>
    internal static OfficeTrueTypeFont? TryLoadFontFamily(string? fontFamily, OfficeFontStyle style,
        out OfficeFontStyle resolvedStyle) => TryLoadFontFamilyForText(
            fontFamily, OfficeFontFaceDescriptor.FromStyle(style), null, out resolvedStyle);

    /// <summary>Resolves a covered installed face using the same ordering as numeric rendering.</summary>
    internal static OfficeTrueTypeFont? TryLoadFontFamilyForText(string? fontFamily, OfficeFontStyle style,
        string? text, out OfficeFontStyle resolvedStyle) => TryLoadFontFamilyForText(
            fontFamily, OfficeFontFaceDescriptor.FromStyle(style), text, out resolvedStyle);

    /// <summary>Checks declared family names without treating a platform substitute as the authored family.</summary>
    internal bool MatchesFamilyName(string family) => HasFamilyKey(NormalizeFontFamilyKey(family));

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
        foreach (string path in CandidateKnownFamilyPaths(key)) yield return path;
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

        // Every parsed face retains its backing array. Read a collection once so inspecting
        // several faces does not retain a complete copy of the same large file per face.
        byte[]? collection = TryReadCollection(path);
        if (collection == null) yield break;
        for (int index = 0; index < count; index++) {
            OfficeTrueTypeFont? face = TryLoad(collection, index);
            if (face != null) yield return face;
        }
    }

    private static byte[]? TryReadCollection(string path) {
        try { return File.ReadAllBytes(path); }
        catch (IOException) { return null; }
        catch (UnauthorizedAccessException) { return null; }
        catch (ArgumentException) { return null; }
        catch (NotSupportedException) { return null; }
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

}
