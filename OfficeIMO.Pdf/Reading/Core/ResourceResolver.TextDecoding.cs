using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static readonly ConditionalWeakTable<Lazy<string?[]?>, Lazy<string?[]>> BuiltinEncodingValues = new();
    private static readonly Lazy<string[]> WinAnsiCodeValues = new(() =>
        Enumerable.Range(0, 256).Select(static code =>
            PdfWinAnsiEncoding.Decode(new[] { (byte)code }, int.MaxValue)).ToArray());

    private static string?[] GetBuiltinEncodingValues(PdfFontResource font) =>
        BuiltinEncodingValues.GetValue(font.BuiltinEncodingSource!, static source =>
            new Lazy<string?[]>(() => {
                string?[]? names = source.Value;
                var values = new string?[256];
                if (names == null) return values;
                for (int code = 0; code < values.Length; code++) {
                    string? name = names[code];
                    if (name != null && TryDecodeGlyphName(name, out string? value))
                        values[code] = value;
                }
                return values;
            })).Value;

    private static System.Func<byte[], string> BuildDecoderForFont(PdfFontResource font, int maxDecodedTextCharacters) {
        if (HasCompositeTextSources(font)) {
            return bytes => DecodeCompositeWithTextSources(font, bytes, maxDecodedTextCharacters, preferToUnicode: true, useWinAnsiFallback: true);
        }
        if (UsesNamedCompositeEncoding(font)) {
            return bytes => DecodeNamedComposite(font, bytes, maxDecodedTextCharacters);
        }
        // Prefer font-specific ToUnicode map when present
        if (font.HasToUnicode && font.CMap is not null) {
            return bytes => DecodeSimpleToUnicode(font, font.CMap, bytes, maxDecodedTextCharacters, allowRawBytes: true);
        }
        var baseDecoder = BuildFontBaseDecoder(font, maxDecodedTextCharacters);
        if (font.Differences is not null && font.Differences.Count > 0) {
            var differences = font.Differences;
            return bytes => DecodeWithDifferences(bytes, differences, baseDecoder, maxDecodedTextCharacters);
        }

        return baseDecoder;
    }

    internal static System.Func<byte[], int, string> CreateBudgetedDecoder(PdfFontResource font) => BuildBudgetedDecoderForFont(font);

    /// <summary>Honors explicit mapping and fallback policy without changing the default decoder.</summary>
    internal static Func<byte[], int, string> CreateBudgetedDecoder(PdfFontResource font,
        bool preferToUnicode, bool useWinAnsiFallback) {
        if (preferToUnicode && useWinAnsiFallback) return BuildBudgetedDecoderForFont(font);
        if (HasCompositeTextSources(font)) {
            return (bytes, maximum) => DecodeCompositeWithTextSources(font, bytes, maximum, preferToUnicode, useWinAnsiFallback);
        }
        if (UsesNamedCompositeEncoding(font)) {
            return (bytes, maximum) => DecodeNamedComposite(font, bytes, maximum, preferToUnicode);
        }
        if (preferToUnicode && font.HasToUnicode && font.CMap is not null) {
            return (bytes, maximum) => DecodeSimpleToUnicode(font, font.CMap, bytes, maximum, allowRawBytes: false);
        }
        Func<byte[], int, string> baseDecoder = BuildBudgetedFontBaseDecoder(font, useWinAnsiFallback);
        return (bytes, maximum) => {
            if (!useWinAnsiFallback && font.Differences != null) {
                foreach (byte code in bytes)
                    if (font.Differences.TryGetValue(code, out string? value) && value.Length == 0)
                        throw new PdfUnsupportedTextMappingException(font);
            }
            bool supportedEncoding = font.HasBuiltinEncoding || font.Encoding is "StandardEncoding" or "MacRomanEncoding" or "SymbolEncoding" or "ZapfDingbatsEncoding" ||
                font.HasDeclaredEncoding && font.Encoding == "WinAnsiEncoding";
            if (!useWinAnsiFallback && !supportedEncoding) {
                var text = new StringBuilder();
                foreach (byte code in bytes) {
                    if (font.Differences is null || !font.Differences.TryGetValue(code, out string? difference) || difference is null) {
                        throw new PdfUnsupportedTextMappingException(font);
                    }
                    if ((long)text.Length + difference.Length > maximum) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximum,
                            (long)text.Length + difference.Length);
                    }
                    text.Append(difference);
                }
                return text.ToString();
            }
            return font.Differences is { Count: > 0 }
                ? DecodeWithDifferences(bytes, font.Differences, input => baseDecoder(input, maximum), maximum)
                : baseDecoder(bytes, maximum);
        };
    }

    /// <summary>Decodes one simple-font code through its encoding and Differences, ignoring ToUnicode.</summary>
    internal static System.Func<byte, string> CreateSimpleEncodingDecoder(PdfFontResource font) {
        System.Func<byte[], int, string> baseDecoder = BuildBudgetedFontBaseDecoder(font);
        IReadOnlyDictionary<int, string>? differences = font.Differences;
        return code => differences != null && differences.TryGetValue(code, out string? difference)
            ? difference
            : baseDecoder(new[] { code }, 1);
    }

    private static System.Func<byte[], int, string> BuildBudgetedDecoderForFont(PdfFontResource font) {
        if (HasCompositeTextSources(font)) {
            return (bytes, maximumCharacters) => DecodeCompositeWithTextSources(font, bytes, maximumCharacters, preferToUnicode: true, useWinAnsiFallback: true);
        }
        if (UsesNamedCompositeEncoding(font)) {
            return (bytes, maximumCharacters) => DecodeNamedComposite(font, bytes, maximumCharacters);
        }
        if (font.HasToUnicode && font.CMap is not null) {
            return (bytes, maximumCharacters) => DecodeSimpleToUnicode(font, font.CMap, bytes, maximumCharacters, allowRawBytes: true);
        }

        if (font.Differences is not null && font.Differences.Count > 0) {
            var differences = font.Differences;
            Func<byte[], int, string> baseDecoder = BuildBudgetedFontBaseDecoder(font);
            return (bytes, maximumCharacters) => DecodeWithDifferences(
                bytes,
                differences,
                input => baseDecoder(input, maximumCharacters),
                maximumCharacters);
        }

        return BuildBudgetedFontBaseDecoder(font);
    }

    /// <summary>
    /// Decodes through ToUnicode. A simple font whose one-byte map leaves codes out reads those codes through its
    /// Differences or declared encoding, as viewers do; codes neither supplies keep the raw fallback or refuse the run.
    /// </summary>
    private static string DecodeSimpleToUnicode(PdfFontResource font, ToUnicodeCMap map, byte[] bytes, int maximum, bool allowRawBytes) {
        if (map.TryMapBytes(bytes, maximum, out string mapped)) return mapped;
        if (map.MaxCodeBytes != 1 || string.Equals(font.FontSubtype, "Type0", StringComparison.Ordinal)) {
            if (allowRawBytes) return map.MapBytes(bytes, maximum);
            throw new PdfUnsupportedTextMappingException(font);
        }
        var text = new StringBuilder(bytes.Length);
        for (int index = 0; index < bytes.Length; index++) {
            string? value = map.TryMapCode(bytes, index, 1, out string code) ? code : GetSimpleEncodingText(font, bytes[index]);
            if (value == null) {
                if (!allowRawBytes) throw new PdfUnsupportedTextMappingException(font);
                value = ((char)bytes[index]).ToString();
            }
            if ((long)text.Length + value.Length > maximum) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximum, (long)text.Length + value.Length);
            }
            text.Append(value);
        }
        return text.ToString();
    }

    // The text a simple font's own encoding gives a code: Differences, a built-in encoding, or a declared base encoding.
    private static string? GetSimpleEncodingText(PdfFontResource font, byte code) {
        string? value;
        if (font.Differences != null && font.Differences.TryGetValue(code, out string? difference)) {
            value = difference;
        } else if (font.HasBuiltinEncoding) {
            value = GetBuiltinEncodingValues(font)[code];
        } else if (font.Encoding is "StandardEncoding" or "MacRomanEncoding" or "SymbolEncoding" or "ZapfDingbatsEncoding" ||
            font.HasDeclaredEncoding && font.Encoding == "WinAnsiEncoding") {
            value = BuildBudgetedBaseEncodingDecoder(font.Encoding)(new[] { code }, 8);
        } else {
            return null;
        }
        return string.IsNullOrEmpty(value) || value![0] < ' ' ? null : value;
    }

    // Named composite CMaps select CIDs, not single-byte WinAnsi characters. Identity
    // encodings, whose code is the CID, decode through the composite text sources.
    // Refuse only when text is shown, preserving unused resources and state-only objects.
    private static bool UsesNamedCompositeEncoding(PdfFontResource font) =>
        string.Equals(font.FontSubtype, "Type0", StringComparison.Ordinal) &&
        !string.Equals(font.Encoding, "Identity-H", StringComparison.Ordinal) &&
        !string.Equals(font.Encoding, "Identity-V", StringComparison.Ordinal);

    private static string DecodeNamedComposite(PdfFontResource font, byte[] bytes, int maximumCharacters, bool preferToUnicode = true) {
        if (bytes.Length == 0) return string.Empty;
        if (preferToUnicode && font.HasToUnicode) {
            if (font.CMap != null && font.CMap.TryMapBytes(bytes, maximumCharacters, out string mapped)) return mapped;
        } else if (font.PredefinedCMap is Lazy<PdfPredefinedCMap> predefined &&
            predefined.Value.TryDecode(bytes, maximumCharacters, out string decoded)) {
            return decoded;
        }
        throw new PdfUnsupportedTextMappingException(font);
    }

    private static Func<byte[], string> BuildFontBaseDecoder(PdfFontResource font, int maximum, bool useWinAnsiFallback = true) {
        Func<byte[], int, string> decoder = BuildBudgetedFontBaseDecoder(font, useWinAnsiFallback);
        return bytes => decoder(bytes, maximum);
    }

    private static Func<byte[], int, string> BuildBudgetedFontBaseDecoder(PdfFontResource font, bool useWinAnsiFallback = true) {
        if (!font.HasBuiltinEncoding) return BuildBudgetedBaseEncodingDecoder(font.Encoding);
        return (bytes, maximum) => {
            if (bytes.Length == 0) return string.Empty;
            string?[] values = GetBuiltinEncodingValues(font);
            string[]? fallback = useWinAnsiFallback ? WinAnsiCodeValues.Value : null;
            if (bytes.Length == 1) {
                string? single = values[bytes[0]] ?? fallback?[bytes[0]];
                if (single == null) throw new PdfUnsupportedTextMappingException(font);
                if (single.Length > maximum)
                    throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximum, single.Length);
                return single;
            }
            var text = new StringBuilder();
            foreach (byte code in bytes) {
                string? value = values[code] ?? fallback?[code];
                if (value == null) throw new PdfUnsupportedTextMappingException(font);
                if ((long)text.Length + value.Length > maximum)
                    throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximum, (long)text.Length + value.Length);
                text.Append(value);
            }
            return text.ToString();
        };
    }

    private static System.Func<byte[], int, string> BuildBudgetedBaseEncodingDecoder(string encoding) {
        if (PdfSymbolicFontEncoding.IsSymbolicEncoding(encoding)) {
            return (bytes, maximumCharacters) => PdfSymbolicFontEncoding.Decode(encoding, bytes, maximumCharacters);
        }

        if (string.Equals(encoding, "StandardEncoding", System.StringComparison.Ordinal)) {
            return static (bytes, maximumCharacters) => PdfStandardEncoding.Decode(bytes, maximumCharacters);
        }

        if (string.Equals(encoding, "MacRomanEncoding", System.StringComparison.Ordinal)) {
            return static (bytes, maximumCharacters) => PdfMacRomanEncoding.Decode(bytes, maximumCharacters);
        }

        return static (bytes, maximumCharacters) => PdfWinAnsiEncoding.Decode(bytes, maximumCharacters);
    }

    private static string DecodeWithDifferences(
        byte[] bytes,
        IReadOnlyDictionary<int, string> differences,
        System.Func<byte[], string> baseDecoder,
        int maxDecodedTextCharacters) {
        if (bytes is null || bytes.Length == 0) return string.Empty;
        if (bytes.LongLength > maxDecodedTextCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxDecodedTextCharacters, bytes.LongLength);
        }
        if (bytes.Length == 1) {
            // The base decoder receives the same one-byte code it would get from the loop below.
            string single = differences.TryGetValue(bytes[0], out string? singleDifference)
                ? singleDifference!
                : baseDecoder(bytes);
            if (single.Length > maxDecodedTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxDecodedTextCharacters, single.Length);
            }
            return single;
        }
        var builder = new System.Text.StringBuilder(bytes.Length);
        byte[]? singleCode = null;
        for (int i = 0; i < bytes.Length; i++) {
            int code = bytes[i];
            string value;
            if (differences.TryGetValue(code, out string? difference)) {
                value = difference!;
            } else {
                // Base decoders read the code without retaining it.
                singleCode ??= new byte[1];
                singleCode[0] = bytes[i];
                value = baseDecoder(singleCode);
            }
            long nextLength = (long)builder.Length + value.Length;
            if (nextLength > maxDecodedTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxDecodedTextCharacters, nextLength);
            }

            builder.Append(value);
        }

        return builder.ToString();
    }

}
