namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static System.Func<byte[], string> BuildDecoderForFont(PdfFontResource font, int maxDecodedTextCharacters) {
        if (UsesNamedCompositeEncoding(font)) {
            return bytes => DecodeNamedComposite(font, bytes, maxDecodedTextCharacters);
        }
        // Prefer font-specific ToUnicode map when present
        if (font.HasToUnicode && font.CMap is not null) return bytes => font.CMap.MapBytes(bytes, maxDecodedTextCharacters);
        var baseDecoder = BuildBaseEncodingDecoder(font.Encoding, maxDecodedTextCharacters);
        if (font.Differences is not null && font.Differences.Count > 0) {
            var differences = font.Differences;
            return bytes => DecodeWithDifferences(bytes, differences, baseDecoder, maxDecodedTextCharacters);
        }

        return baseDecoder;
    }

    internal static System.Func<byte[], int, string> CreateBudgetedDecoder(PdfFontResource font) => BuildBudgetedDecoderForFont(font);

    /// <summary>Decodes one simple-font code through its encoding and Differences, ignoring ToUnicode.</summary>
    internal static System.Func<byte, string> CreateSimpleEncodingDecoder(PdfFontResource font) {
        System.Func<byte[], int, string> baseDecoder = BuildBudgetedBaseEncodingDecoder(font.Encoding);
        IReadOnlyDictionary<int, string>? differences = font.Differences;
        return code => differences != null && differences.TryGetValue(code, out string? difference)
            ? difference
            : baseDecoder(new[] { code }, 1);
    }

    private static System.Func<byte[], int, string> BuildBudgetedDecoderForFont(PdfFontResource font) {
        if (UsesNamedCompositeEncoding(font)) {
            return (bytes, maximumCharacters) => DecodeNamedComposite(font, bytes, maximumCharacters);
        }
        if (font.HasToUnicode && font.CMap is not null) {
            return (bytes, maximumCharacters) => font.CMap.MapBytes(bytes, maximumCharacters);
        }

        if (font.Differences is not null && font.Differences.Count > 0) {
            var differences = font.Differences;
            return (bytes, maximumCharacters) => DecodeWithDifferences(
                bytes,
                differences,
                BuildBaseEncodingDecoder(font.Encoding, maximumCharacters),
                maximumCharacters);
        }

        return BuildBudgetedBaseEncodingDecoder(font.Encoding);
    }

    // Named composite CMaps select CIDs, not single-byte WinAnsi characters. Identity
    // encodings retain the existing raw-CID rendering and editability paths.
    // Refuse only when text is shown, preserving unused resources and state-only objects.
    private static bool UsesNamedCompositeEncoding(PdfFontResource font) =>
        string.Equals(font.FontSubtype, "Type0", StringComparison.Ordinal) &&
        !string.Equals(font.Encoding, "Identity-H", StringComparison.Ordinal) &&
        !string.Equals(font.Encoding, "Identity-V", StringComparison.Ordinal);

    private static string DecodeNamedComposite(PdfFontResource font, byte[] bytes, int maximumCharacters) {
        if (bytes.Length == 0) return string.Empty;
        if (font.HasToUnicode) {
            if (font.CMap != null && font.CMap.TryMapBytes(bytes, maximumCharacters, out string mapped)) return mapped;
        } else if (font.PredefinedCMap is Lazy<PdfPredefinedCMap> predefined &&
            predefined.Value.TryDecode(bytes, maximumCharacters, out string decoded)) {
            return decoded;
        }
        throw new PdfUnsupportedTextMappingException(font);
    }

    private static System.Func<byte[], int, string> BuildBudgetedBaseEncodingDecoder(string encoding) {
        if (string.Equals(encoding, "StandardEncoding", System.StringComparison.Ordinal)) {
            return static (bytes, maximumCharacters) => PdfStandardEncoding.Decode(bytes, maximumCharacters);
        }

        if (string.Equals(encoding, "MacRomanEncoding", System.StringComparison.Ordinal)) {
            return static (bytes, maximumCharacters) => PdfMacRomanEncoding.Decode(bytes, maximumCharacters);
        }

        return static (bytes, maximumCharacters) => PdfWinAnsiEncoding.Decode(bytes, maximumCharacters);
    }

    private static System.Func<byte[], string> BuildBaseEncodingDecoder(string encoding, int maxDecodedTextCharacters) {
        if (string.Equals(encoding, "StandardEncoding", System.StringComparison.Ordinal)) {
            return bytes => PdfStandardEncoding.Decode(bytes, maxDecodedTextCharacters);
        }

        if (string.Equals(encoding, "MacRomanEncoding", System.StringComparison.Ordinal)) {
            return bytes => PdfMacRomanEncoding.Decode(bytes, maxDecodedTextCharacters);
        }

        return bytes => PdfWinAnsiEncoding.Decode(bytes, maxDecodedTextCharacters);
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
        var builder = new System.Text.StringBuilder(bytes.Length);
        for (int i = 0; i < bytes.Length; i++) {
            int code = bytes[i];
            string value = differences.TryGetValue(code, out string? difference)
                ? difference!
                : baseDecoder(new[] { bytes[i] });
            long nextLength = (long)builder.Length + value.Length;
            if (nextLength > maxDecodedTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxDecodedTextCharacters, nextLength);
            }

            builder.Append(value);
        }

        return builder.ToString();
    }

}
