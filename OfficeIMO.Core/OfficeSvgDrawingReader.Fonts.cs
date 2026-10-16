using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumSvgEmbeddedFontFaces = 128;
    private const int MaximumSvgDecodedFontBytes = 8 * 1024 * 1024;
    private const int MaximumSvgRetainedFontBytes = 32 * 1024 * 1024;
    private const int MaximumSvgFontLoadWorkBytes = 128 * 1024 * 1024;

    /// <summary>Per-document font limits include failed loads and repeated registrations, not just final faces.</summary>
    private sealed class SvgEmbeddedFontImports {
        internal SvgEmbeddedFontImports(OfficeFontFaceCollection fonts, CancellationToken cancellationToken) {
            Fonts = fonts;
            CancellationToken = cancellationToken;
        }

        internal OfficeFontFaceCollection Fonts { get; }
        internal CancellationToken CancellationToken { get; }
        internal int FaceCount;
        internal int RemainingRetainedBytes = MaximumSvgRetainedFontBytes;
        internal int RemainingLoadWorkBytes = MaximumSvgFontLoadWorkBytes;
    }

    /// <summary>Imports only embedded fonts; no URL resolver, installed-font loader or caller callback is invoked.</summary>
    private static bool TryImportSvgFontFace(string body, SvgEmbeddedFontImports imports, ref int declarationCount) {
        imports.CancellationToken.ThrowIfCancellationRequested();
        if (++imports.FaceCount > MaximumSvgEmbeddedFontFaces) return false;
        IReadOnlyList<SvgCssDeclaration> declarations = ParseSvgCssDeclarations(body, ref declarationCount, out bool complete);
        if (!complete) return false;
        var descriptors = new Dictionary<string, string>(StringComparer.Ordinal);
        // The common declaration reader tolerates malformed declarations. Font loading must not
        // silently bless a partial descriptor block as complete CSS during content inspection.
        if (SplitSvgCssTopLevel(body, ';').Count(item => !string.IsNullOrWhiteSpace(item)) != declarations.Count) return false;
        foreach (SvgCssDeclaration declaration in declarations) {
            if (declaration.Important) return false;
            switch (declaration.Name) {
                case "font-family": case "src": case "font-weight": case "font-stretch": case "font-style": case "unicode-range":
                    descriptors[declaration.Name] = declaration.Value;
                    break;
                default: return false;
            }
        }
        if (!descriptors.TryGetValue("font-family", out string? familyValue) ||
            !TryReadSvgFontFamily(familyValue, out string family) ||
            !descriptors.TryGetValue("src", out string? sources)) return false;
        string weightValue = GetSvgFontDescriptor(descriptors, "font-weight", "normal");
        string stretchValue = GetSvgFontDescriptor(descriptors, "font-stretch", "normal");
        string styleValue = GetSvgFontDescriptor(descriptors, "font-style", "normal");
        if (IsRelativeSvgFontDescriptor(weightValue) || IsRelativeSvgFontDescriptor(stretchValue) || IsRelativeSvgFontDescriptor(styleValue) ||
            !OfficeFontFaceCssParser.TryWeight(weightValue, 400, out int weight) ||
            !OfficeFontFaceCssParser.TryStretch(stretchValue, 100D, out double stretch) ||
            !OfficeFontFaceCssParser.TrySlant(styleValue, OfficeFontFaceDescriptor.Regular, out OfficeFontSlant slant, out double angle)) return false;
        var descriptor = new OfficeFontFaceDescriptor(weight, stretch, slant, angle);
        OfficeFontUnicodeRangeSet? ranges = OfficeFontUnicodeRangeSet.All;
        if (descriptors.TryGetValue("unicode-range", out string? rangeValue) &&
            !OfficeFontUnicodeRangeSet.TryParseCss(rangeValue, out ranges)) return false;

        bool loaded = false;
        foreach (string source in SplitSvgCssTopLevel(sources, ',')) {
            imports.CancellationToken.ThrowIfCancellationRequested();
            if (!TryReadSvgEmbeddedFontSource(source.Trim(), out string dataUri)) {
                complete = false;
                continue;
            }
            int limit = Math.Min(MaximumSvgDecodedFontBytes, Math.Min(imports.RemainingRetainedBytes, imports.RemainingLoadWorkBytes));
            if (limit <= 0 || !TryDecodeSvgFontDataUri(dataUri, limit, out byte[] data)) {
                complete = false;
                continue;
            }
            if (loaded) continue;
            imports.CancellationToken.ThrowIfCancellationRequested();
            // Use only the first-party bounded loader for untrusted document bytes. Providers and
            // variation resolvers belong to the caller's explicit registrations, not SVG content.
            var accepted = new OfficeFontFaceCollection();
            loaded = accepted.TryAddBounded(family, data, descriptor, ranges, limit, out int retainedBytes, out _, applyDescriptorWeight: true);
            imports.CancellationToken.ThrowIfCancellationRequested();
            // A failed decoder may have allocated its entire allowance before rejecting input.
            imports.RemainingLoadWorkBytes -= loaded ? retainedBytes : limit;
            if (loaded) {
                imports.RemainingRetainedBytes -= retainedBytes;
                imports.Fonts.AddRange(accepted);
            } else complete = false;
        }
        return loaded && complete;
    }

    private static string GetSvgFontDescriptor(IReadOnlyDictionary<string, string> descriptors, string name, string fallback) =>
        descriptors.TryGetValue(name, out string? value) ? value : fallback;

    private static bool IsRelativeSvgFontDescriptor(string value) =>
        IsSvgCssWideKeyword(value) || value.Equals("bolder", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("lighter", StringComparison.OrdinalIgnoreCase) || value.Equals("wider", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("narrower", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadSvgFontFamily(string value, out string family) {
        family = string.Empty;
        string normalized = value.Trim();
        if (normalized.Length == 0) return false;
        if (normalized[0] is '\'' or '"') {
            int cursor = 0;
            if (!TryReadSvgFontCssString(normalized, ref cursor, out family) || cursor != normalized.Length) return false;
        } else {
            if (normalized.Length > OfficeFontFamilyParser.DefaultMaximumFamilyNameLength) return false;
            if (IsSvgCssWideKeyword(normalized) || normalized.Equals("default", StringComparison.OrdinalIgnoreCase)) return false;
            foreach (string part in normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                if (!IsSupportedSvgSelectorIdentifier(part)) return false;
            }
            List<string> families = OfficeFontFamilyParser.Parse(normalized);
            if (families.Count != 1) return false;
            family = families[0];
        }
        return !string.IsNullOrWhiteSpace(family) && family.Length <= OfficeFontFamilyParser.DefaultMaximumFamilyNameLength;
    }

    /// <summary>Reads CSS quoted strings using the shared escape decoder, without accepting trailing syntax.</summary>
    private static bool TryReadSvgFontCssString(string value, ref int cursor, out string text) {
        text = string.Empty;
        if (cursor >= value.Length || value[cursor] is not ('\'' or '"')) return false;
        char quote = value[cursor++];
        var result = new StringBuilder();
        while (cursor < value.Length) {
            char current = value[cursor];
            if (current == quote) {
                cursor++;
                text = result.ToString();
                return true;
            }
            if (current is '\r' or '\n' or '\f' || !TryReadCssCharacter(value, ref cursor, out char decoded) || decoded == '\0') return false;
            result.Append(decoded);
        }
        return false;
    }

    private static bool TryReadSvgEmbeddedFontSource(string source, out string dataUri) {
        dataUri = string.Empty;
        int cursor = 0;
        if (!TryReadCssIdentifier(source, ref cursor, "url") || cursor >= source.Length || source[cursor++] != '(') return false;
        SkipSvgFontCssWhitespace(source, ref cursor);
        if (cursor < source.Length && source[cursor] is '\'' or '"') {
            if (!TryReadSvgFontCssString(source, ref cursor, out dataUri)) return false;
            SkipSvgFontCssWhitespace(source, ref cursor);
        } else {
            int start = cursor;
            while (cursor < source.Length && source[cursor] != ')' && !IsSvgCssWhitespace(source[cursor])) {
                if (source[cursor] is '(' or '\'' or '"' or '\\' || char.IsControl(source[cursor])) return false;
                cursor++;
            }
            dataUri = source.Substring(start, cursor - start);
            SkipSvgFontCssWhitespace(source, ref cursor);
        }
        if (cursor >= source.Length || source[cursor++] != ')' || !dataUri.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return false;
        SkipSvgFontCssWhitespace(source, ref cursor);
        if (cursor == source.Length) return true;
        if (!TryReadCssIdentifier(source, ref cursor, "format") || cursor >= source.Length || source[cursor++] != '(') return false;
        SkipSvgFontCssWhitespace(source, ref cursor);
        if (!TryReadSvgFontCssString(source, ref cursor, out string format)) return false;
        if (!(format.Equals("truetype", StringComparison.OrdinalIgnoreCase) || format.Equals("opentype", StringComparison.OrdinalIgnoreCase) ||
              format.Equals("woff", StringComparison.OrdinalIgnoreCase) || format.Equals("woff2", StringComparison.OrdinalIgnoreCase))) return false;
        SkipSvgFontCssWhitespace(source, ref cursor);
        if (cursor >= source.Length || source[cursor++] != ')') return false;
        SkipSvgFontCssWhitespace(source, ref cursor);
        return cursor == source.Length;
    }

    private static void SkipSvgFontCssWhitespace(string value, ref int cursor) {
        while (cursor < value.Length && IsSvgCssWhitespace(value[cursor])) cursor++;
    }

    private static bool TryDecodeSvgFontDataUri(string source, int maximumBytes, out byte[] data) {
        data = Array.Empty<byte>();
        int comma = source.IndexOf(',');
        if (comma < 0) return false;
        string metadata = source.Substring(5, comma - 5);
        if (!(metadata.Equals("font/ttf;base64", StringComparison.OrdinalIgnoreCase) ||
              metadata.Equals("font/otf;base64", StringComparison.OrdinalIgnoreCase) ||
              metadata.Equals("font/woff;base64", StringComparison.OrdinalIgnoreCase) ||
              metadata.Equals("font/woff2;base64", StringComparison.OrdinalIgnoreCase))) return false;
        long payloadLength = source.Length - comma - 1L;
        if ((payloadLength + 3L) / 4L * 3L - 2L > maximumBytes) return false;
        try {
            data = Convert.FromBase64String(source.Substring(comma + 1));
        } catch (FormatException) {
            return false;
        }
        if (data.Length == 0 || data.Length > maximumBytes) return false;
        OfficeFontContainerFormat format = OfficeFontContainerDecoder.Detect(data);
        return metadata.StartsWith("font/woff2;", StringComparison.OrdinalIgnoreCase) ? format == OfficeFontContainerFormat.Woff2
            : metadata.StartsWith("font/woff;", StringComparison.OrdinalIgnoreCase) ? format == OfficeFontContainerFormat.Woff
            : format == OfficeFontContainerFormat.OpenType;
    }
}
