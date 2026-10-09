using System.Collections;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>
/// Decides whether a document needs the Japanese, Arabic and symbol fallback fonts. They live in their own
/// assembly so the browser engine downloads them only for documents whose text Carlito can't draw.
/// </summary>
internal static partial class BrowserFallbackFonts {
    internal const string AssemblyName = "OfficeIMO.Web.Fonts.Fallback";
    internal const string AssetName = AssemblyName + ".wasm";

    /// <summary>Decompressed XML read before giving up and assuming the fallback fonts are needed.</summary>
    private const long MaximumScannedBytes = 128L * 1024 * 1024;

    private static readonly Lazy<CarlitoCoverage> Coverage = new(() => new CarlitoCoverage(BrowserPortablePdfProfile.CarlitoRegularBytes), isThreadSafe: true);

    /// <summary>The fallback font assembly once it is in the process. In the browser that means the worker has lazy-loaded it.</summary>
    internal static Assembly? FindAssembly() {
        foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies()) {
            if (string.Equals(assembly.GetName().Name, AssemblyName, StringComparison.Ordinal)) return assembly;
        }
        if (OperatingSystem.IsBrowser()) return null;
        try {
            return Assembly.Load(new AssemblyName(AssemblyName));
        } catch (Exception error) when (error is FileNotFoundException or FileLoadException or BadImageFormatException) {
            return null;
        }
    }

    internal static bool IsAvailable => FindAssembly() != null;

    /// <summary>True when a file's text or font choices need a fallback font. Unknown or unreadable input counts as needing it.</summary>
    internal static bool Needed(byte[] bytes, string extension) {
        try {
            return extension.ToLowerInvariant() switch {
                ".docx" or ".docm" or ".dotx" or ".dotm" or
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or
                ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => PackageNeeds(bytes),
                ".html" or ".htm" or ".xhtml" or ".md" or ".markdown" or ".txt" => Needed(Decode(bytes)),
                _ => true
            };
        } catch (Exception error) when (error is InvalidDataException or IOException or DecoderFallbackException) {
            // The conversion will report the damage; loading fonts first changes nothing.
            return true;
        }
    }

    /// <summary>True when text contains a character Carlito lacks, an entity that may produce one, or names a symbol font.</summary>
    internal static bool Needed(string text) {
        CarlitoCoverage coverage = Coverage.Value;
        for (int index = 0; index < text.Length; index++) {
            char character = text[index];
            if (character < 0x80) continue;
            int scalar = character;
            if (char.IsHighSurrogate(character) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                scalar = char.ConvertToUtf32(character, text[++index]);
            }
            if (!Ignorable(scalar) && !coverage.Contains(scalar)) return true;
        }
        foreach (Match entity in CharacterReference().Matches(text)) {
            string value = entity.Groups[1].Value;
            int scalar;
            if (value[0] == '#') {
                bool hex = value.Length > 1 && (value[1] == 'x' || value[1] == 'X');
                if (!int.TryParse(hex ? value.AsSpan(2) : value.AsSpan(1), hex ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out scalar)) continue;
                if (!Ignorable(scalar) && !coverage.Contains(scalar)) return true;
            } else if (!SafeNamedEntities.Contains(value)) {
                return true;
            }
        }
        return SymbolFontName().IsMatch(text);
    }

    private static readonly HashSet<string> SafeNamedEntities = new(StringComparer.Ordinal) {
        "amp", "lt", "gt", "quot", "apos", "nbsp", "copy", "reg", "trade", "ndash", "mdash", "hellip",
        "lsquo", "rsquo", "ldquo", "rdquo", "bull", "middot", "euro", "pound", "yen", "cent", "deg", "times", "divide", "shy"
    };

    [GeneratedRegex(@"&(#[xX]?[0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});", RegexOptions.CultureInvariant)]
    private static partial Regex CharacterReference();

    /// <summary>Font names the pack maps to Noto Sans Symbols 2 (see the profile's symbol aliases and font-pack.json).</summary>
    [GeneratedRegex(@"\b(?:Symbol|ZapfDingbats|Zapf\s+Dingbats)\b", RegexOptions.CultureInvariant)]
    private static partial Regex SymbolFontName();

    /// <summary>Layout controls and invisible characters that need no glyph.</summary>
    private static bool Ignorable(int scalar) =>
        scalar is < 0x20 or (>= 0x7F and <= 0x9F) or 0xAD or (>= 0x200B and <= 0x200F) or (>= 0x2028 and <= 0x202E) or
            (>= 0x2060 and <= 0x206F) or 0xFEFF or (>= 0xFE00 and <= 0xFE0F) or (>= 0xFFF0 and <= 0xFFFF);

    private static bool PackageNeeds(byte[] bytes) {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        long budget = MaximumScannedBytes;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (!Scanned(entry.FullName)) continue;
            budget -= entry.Length;
            if (budget < 0) return true;
            using Stream stream = entry.Open();
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            if (Needed(reader.ReadToEnd())) return true;
        }
        return false;
    }

    /// <summary>
    /// Parts whose text or font choices are drawn. Themes and font tables list a native-script font name for every
    /// script (Japanese, Arabic, Thai...) in every document, and are used only when the text itself needs them.
    /// </summary>
    private static bool Scanned(string name) {
        if (!name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return false;
        if (name.StartsWith("docProps/", StringComparison.OrdinalIgnoreCase) ||
            name.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("/_rels/", StringComparison.Ordinal) || name.StartsWith("_rels/", StringComparison.Ordinal) ||
            name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) return false;
        string file = name.Substring(name.LastIndexOf('/') + 1);
        return !(name.Contains("/theme/", StringComparison.OrdinalIgnoreCase) ||
                 file.Equals("fontTable.xml", StringComparison.OrdinalIgnoreCase) ||
                 file.Equals("settings.xml", StringComparison.OrdinalIgnoreCase) ||
                 file.Equals("webSettings.xml", StringComparison.OrdinalIgnoreCase));
    }

    private static string Decode(byte[] bytes) {
        using var reader = new StreamReader(new MemoryStream(bytes, writable: false), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    /// <summary>The Unicode scalars Carlito Regular maps to glyphs, read from its cmap (formats 4 and 12).</summary>
    private sealed class CarlitoCoverage {
        private readonly BitArray _basic = new(0x10000);
        private readonly List<(int Start, int End)> _supplementary = [];

        internal CarlitoCoverage(byte[] font) {
            int tables = U16(font, 4);
            for (int table = 0; table < tables; table++) {
                int record = 12 + table * 16;
                if (Encoding.ASCII.GetString(font, record, 4) != "cmap") continue;
                ReadCmap(font, (int)U32(font, record + 8));
                return;
            }
            throw new InvalidDataException("The default font has no character map.");
        }

        internal bool Contains(int scalar) {
            if (scalar < 0x10000) return _basic[scalar];
            foreach ((int start, int end) in _supplementary) {
                if (scalar >= start && scalar <= end) return true;
            }
            return false;
        }

        private void ReadCmap(byte[] font, int cmap) {
            int count = U16(font, cmap + 2);
            for (int index = 0; index < count; index++) {
                int record = cmap + 4 + index * 8;
                int platform = U16(font, record), encoding = U16(font, record + 2);
                if (platform != 0 && !(platform == 3 && encoding is 1 or 10)) continue;
                int subtable = cmap + (int)U32(font, record + 4);
                switch (U16(font, subtable)) {
                    case 4: ReadFormat4(font, subtable); break;
                    case 12: ReadFormat12(font, subtable); break;
                }
            }
        }

        private void ReadFormat4(byte[] font, int table) {
            int segments = U16(font, table + 6) / 2;
            int ends = table + 14, starts = ends + segments * 2 + 2, deltas = starts + segments * 2, offsets = deltas + segments * 2;
            for (int segment = 0; segment < segments; segment++) {
                int end = U16(font, ends + segment * 2), start = U16(font, starts + segment * 2);
                int delta = U16(font, deltas + segment * 2), rangeOffset = U16(font, offsets + segment * 2);
                for (int code = start; code <= end && code != 0xFFFF; code++) {
                    int glyph;
                    if (rangeOffset == 0) {
                        glyph = (code + delta) & 0xFFFF;
                    } else {
                        int at = offsets + segment * 2 + rangeOffset + (code - start) * 2;
                        glyph = at + 1 < font.Length ? U16(font, at) : 0;
                        if (glyph != 0) glyph = (glyph + delta) & 0xFFFF;
                    }
                    if (glyph != 0) _basic[code] = true;
                }
            }
        }

        private void ReadFormat12(byte[] font, int table) {
            long groups = U32(font, table + 12);
            for (long group = 0; group < groups; group++) {
                int at = table + 16 + (int)group * 12;
                int start = (int)U32(font, at), end = (int)U32(font, at + 4);
                for (int code = start; code <= Math.Min(end, 0xFFFF); code++) _basic[code] = true;
                if (end >= 0x10000) _supplementary.Add((Math.Max(start, 0x10000), end));
            }
        }

        private static int U16(byte[] data, int offset) => data[offset] << 8 | data[offset + 1];
        private static uint U32(byte[] data, int offset) => (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);
    }
}
