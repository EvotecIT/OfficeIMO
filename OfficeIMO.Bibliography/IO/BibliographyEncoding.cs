namespace OfficeIMO.Bibliography;

internal static class BibliographyEncoding {
    private const int EncodingCharacterChunkSize = 4096;

    internal static Encoding Detect(byte[] bytes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return new UTF8Encoding(true, true);
        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00) return new UTF32Encoding(false, true, true);
        if (bytes.Length >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF) return new UTF32Encoding(true, true, true);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return new UnicodeEncoding(false, true, true);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return new UnicodeEncoding(true, true, true);
        if (StartsWithXmlMarkup(bytes, 4, true, cancellationToken)) return new UTF32Encoding(true, false, true);
        if (StartsWithXmlMarkup(bytes, 4, false, cancellationToken)) return new UTF32Encoding(false, false, true);
        if (StartsWithXmlMarkup(bytes, 2, true, cancellationToken)) return new UnicodeEncoding(true, false, true);
        if (StartsWithXmlMarkup(bytes, 2, false, cancellationToken)) return new UnicodeEncoding(false, false, true);

        var fallback = new UTF8Encoding(false, true);
        return ResolveXmlDeclaration(bytes, fallback, cancellationToken);
    }

    private static Encoding ResolveXmlDeclaration(byte[] bytes, Encoding fallback, CancellationToken cancellationToken) {
        int declarationStart = 0;
        while (declarationStart < bytes.Length && IsAsciiXmlWhitespace(bytes[declarationStart])) {
            if ((declarationStart & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            declarationStart++;
        }
        if (!MatchesAscii(bytes, declarationStart, "<?xml")) return fallback;

        int declarationEnd = -1;
        for (int index = declarationStart + 5; index + 1 < bytes.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (bytes[index] == (byte)'?' && bytes[index + 1] == (byte)'>') { declarationEnd = index; break; }
        }
        if (declarationEnd < 0) return fallback;

        int encodingStart = IndexOfAscii(bytes, declarationStart + 5, declarationEnd, "encoding", cancellationToken);
        if (encodingStart < 0) return fallback;
        int equals = encodingStart + 8;
        while (equals < declarationEnd && IsAsciiXmlWhitespace(bytes[equals])) {
            if ((equals & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            equals++;
        }
        if (equals >= declarationEnd || bytes[equals] != (byte)'=') return fallback;
        int quote = equals + 1;
        while (quote < declarationEnd && IsAsciiXmlWhitespace(bytes[quote])) {
            if ((quote & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            quote++;
        }
        if (quote >= declarationEnd || bytes[quote] != (byte)'\'' && bytes[quote] != (byte)'"') return fallback;
        byte delimiter = bytes[quote++];
        int valueEnd = quote;
        while (valueEnd < declarationEnd && bytes[valueEnd] != delimiter) {
            if ((valueEnd & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            valueEnd++;
        }
        if (valueEnd <= quote || valueEnd >= declarationEnd) return fallback;
        cancellationToken.ThrowIfCancellationRequested();
        return ResolveEncodingName(Encoding.ASCII.GetString(bytes, quote, valueEnd - quote), fallback);
    }

    internal static Encoding ResolveXmlDeclaration(string source, Encoding fallback, CancellationToken cancellationToken) {
        int declarationStart = 0;
        while (declarationStart < source.Length && (char.IsWhiteSpace(source[declarationStart]) || source[declarationStart] == '\uFEFF')) {
            if ((declarationStart & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            declarationStart++;
        }
        if (!MatchesText(source, declarationStart, "<?xml")) return fallback;
        int declarationEnd = IndexOfText(source, declarationStart + 5, source.Length, "?>", cancellationToken, caseInsensitive: false);
        if (declarationEnd < 0) return fallback;
        int encodingStart = IndexOfText(source, declarationStart + 5, declarationEnd, "encoding", cancellationToken, caseInsensitive: true);
        if (encodingStart < 0) return fallback;
        int equals = encodingStart + 8;
        while (equals < declarationEnd && char.IsWhiteSpace(source[equals])) {
            if ((equals & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            equals++;
        }
        if (equals >= declarationEnd || source[equals] != '=') return fallback;
        int quote = equals + 1;
        while (quote < declarationEnd && char.IsWhiteSpace(source[quote])) {
            if ((quote & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            quote++;
        }
        if (quote >= declarationEnd || source[quote] != '\'' && source[quote] != '"') return fallback;
        char delimiter = source[quote++];
        int valueEnd = quote;
        while (valueEnd < declarationEnd && source[valueEnd] != delimiter) {
            if ((valueEnd & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            valueEnd++;
        }
        if (valueEnd <= quote || valueEnd >= declarationEnd) return fallback;
        cancellationToken.ThrowIfCancellationRequested();
        return ResolveEncodingName(source.Substring(quote, valueEnd - quote), fallback);
    }

    internal static string DecodeBounded(byte[] bytes, Encoding encoding, int maximumCharacters, CancellationToken cancellationToken) {
        byte[] preamble = encoding.GetPreamble();
        int offset = HasPreamble(bytes, preamble) ? preamble.Length : 0;
        int end = bytes.Length;
        var decoder = encoding.GetDecoder();
        var characters = new char[4096];
        var builder = new StringBuilder(Math.Min(maximumCharacters, Math.Min(end - offset, characters.Length)));
        bool completed = false;
        while (!completed) {
            cancellationToken.ThrowIfCancellationRequested();
            decoder.Convert(bytes, offset, end - offset, characters, 0, characters.Length, true, out int bytesUsed, out int charactersUsed, out completed);
            if (charactersUsed > maximumCharacters - builder.Length)
                throw new InvalidDataException($"Bibliography input exceeds the configured {maximumCharacters} character limit.");
            builder.Append(characters, 0, charactersUsed);
            offset += bytesUsed;
            if (!completed && bytesUsed == 0 && charactersUsed == 0) throw new InvalidDataException("Bibliography input could not be decoded within the configured character limit.");
        }
        return builder.ToString();
    }

    internal static bool CanEncode(string value, Encoding encoding, CancellationToken cancellationToken) {
        var strictEncoding = (Encoding)encoding.Clone();
        strictEncoding.EncoderFallback = EncoderFallback.ExceptionFallback;
        try {
            ConvertChunks(value, strictEncoding, null, 0, cancellationToken);
            return true;
        } catch (EncoderFallbackException) {
            return false;
        }
    }

    internal static byte[] Encode(string value, Encoding encoding, CancellationToken cancellationToken) {
        var outputEncoding = (Encoding)encoding.Clone();
        outputEncoding.EncoderFallback = EncoderFallback.ReplacementFallback;
        byte[] preamble = outputEncoding.GetPreamble();
        int contentLength = ConvertChunks(value, outputEncoding, null, 0, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var result = new byte[checked(preamble.Length + contentLength)];
        if (preamble.Length > 0) Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length);
        int written = ConvertChunks(value, outputEncoding, result, preamble.Length, cancellationToken);
        if (written != contentLength) throw new InvalidDataException("Bibliography output encoding produced an inconsistent byte count.");
        return result;
    }

    internal static byte[] CloneBytes(byte[] source, CancellationToken cancellationToken) {
        var result = new byte[source.Length];
        for (int offset = 0; offset < source.Length; offset += 1024 * 1024) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(1024 * 1024, source.Length - offset);
            Buffer.BlockCopy(source, offset, result, offset, count);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    private static int ConvertChunks(string value, Encoding encoding, byte[]? destination, int destinationOffset, CancellationToken cancellationToken) {
        Encoder encoder = encoding.GetEncoder();
        var characters = new char[EncodingCharacterChunkSize];
        byte[]? scratch = destination == null ? new byte[encoding.GetMaxByteCount(characters.Length)] : null;
        int sourceOffset = 0;
        int totalBytes = 0;
        while (sourceOffset < value.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int characterCount = Math.Min(characters.Length, value.Length - sourceOffset);
            value.CopyTo(sourceOffset, characters, 0, characterCount);
            bool flush = sourceOffset + characterCount == value.Length;
            int characterOffset = 0;
            bool completed = false;
            while (!completed) {
                cancellationToken.ThrowIfCancellationRequested();
                byte[] bytes = destination ?? scratch!;
                int byteOffset = destination == null ? 0 : destinationOffset + totalBytes;
                int byteCount = destination == null ? bytes.Length : bytes.Length - byteOffset;
                encoder.Convert(characters, characterOffset, characterCount - characterOffset, bytes, byteOffset, byteCount, flush, out int charactersUsed, out int bytesUsed, out completed);
                if (!completed && charactersUsed == 0 && bytesUsed == 0) throw new InvalidDataException("Bibliography output could not be encoded incrementally.");
                characterOffset += charactersUsed;
                totalBytes = checked(totalBytes + bytesUsed);
            }
            sourceOffset += characterCount;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return totalBytes;
    }

    private static bool HasPreamble(byte[] bytes, byte[] preamble) {
        if (preamble.Length == 0 || bytes.Length < preamble.Length) return false;
        for (int index = 0; index < preamble.Length; index++) if (bytes[index] != preamble[index]) return false;
        return true;
    }

    private static bool StartsWithXmlMarkup(byte[] bytes, int width, bool bigEndian, CancellationToken cancellationToken) {
        int maximum = bytes.Length - bytes.Length % width;
        for (int offset = 0; offset < maximum; offset += width) {
            if ((offset & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            uint value = 0;
            for (int index = 0; index < width; index++) {
                int sourceIndex = bigEndian ? offset + index : offset + width - index - 1;
                value = (value << 8) | bytes[sourceIndex];
            }
            if (value == '<') return true;
            if (value != ' ' && value != '\t' && value != '\r' && value != '\n' && value != 0xFEFF) return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }

    private static int IndexOfAscii(byte[] bytes, int start, int end, string value, CancellationToken cancellationToken) {
        int maximum = end - value.Length;
        for (int index = start; index <= maximum; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (MatchesAscii(bytes, index, value)) return index;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return -1;
    }

    private static bool MatchesAscii(byte[] bytes, int start, string value) {
        if (start < 0 || start > bytes.Length - value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            byte current = bytes[start + index];
            char expected = value[index];
            if (current >= (byte)'A' && current <= (byte)'Z') current = (byte)(current + ((byte)'a' - (byte)'A'));
            if (expected >= 'A' && expected <= 'Z') expected = (char)(expected + ('a' - 'A'));
            if (current != (byte)expected) return false;
        }
        return true;
    }

    private static int IndexOfText(string source, int start, int end, string value, CancellationToken cancellationToken, bool caseInsensitive) {
        int maximum = end - value.Length;
        for (int index = start; index <= maximum; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (MatchesText(source, index, value, caseInsensitive)) return index;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return -1;
    }

    private static bool MatchesText(string source, int start, string value, bool caseInsensitive = true) {
        if (start < 0 || start > source.Length - value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            char current = source[start + index];
            char expected = value[index];
            if (caseInsensitive) { current = char.ToLowerInvariant(current); expected = char.ToLowerInvariant(expected); }
            if (current != expected) return false;
        }
        return true;
    }

    private static bool IsAsciiXmlWhitespace(byte value) => value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' || value == (byte)'\n';

    private static Encoding ResolveEncodingName(string name, Encoding fallback) {
        try {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(name, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        } catch (ArgumentException) {
            return fallback;
        }
    }
}
