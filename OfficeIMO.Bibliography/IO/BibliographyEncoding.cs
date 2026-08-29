namespace OfficeIMO.Bibliography;

internal static class BibliographyEncoding {
    internal static Encoding Detect(byte[] bytes) {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return new UTF8Encoding(true, true);
        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00) return new UTF32Encoding(false, true, true);
        if (bytes.Length >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF) return new UTF32Encoding(true, true, true);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return new UnicodeEncoding(false, true, true);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return new UnicodeEncoding(true, true, true);
        if (bytes.Length >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0x00 && bytes[3] == 0x3C) return new UTF32Encoding(true, false, true);
        if (bytes.Length >= 4 && bytes[0] == 0x3C && bytes[1] == 0x00 && bytes[2] == 0x00 && bytes[3] == 0x00) return new UTF32Encoding(false, false, true);
        if (bytes.Length >= 4 && bytes[0] == 0x00 && bytes[1] == 0x3C && bytes[2] == 0x00 && bytes[3] == 0x3F) return new UnicodeEncoding(true, false, true);
        if (bytes.Length >= 4 && bytes[0] == 0x3C && bytes[1] == 0x00 && bytes[2] == 0x3F && bytes[3] == 0x00) return new UnicodeEncoding(false, false, true);

        var fallback = new UTF8Encoding(false, true);
        string prefix = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 4096));
        return ResolveXmlDeclaration(prefix, fallback);
    }

    internal static Encoding ResolveXmlDeclaration(string source, Encoding fallback) {
        int declarationStart = source.IndexOf("<?xml", StringComparison.OrdinalIgnoreCase);
        if (declarationStart < 0 || source.Substring(0, declarationStart).Any(character => !char.IsWhiteSpace(character) && character != '\uFEFF')) return fallback;
        int declarationEnd = source.IndexOf("?>", declarationStart, StringComparison.Ordinal);
        if (declarationEnd < 0) return fallback;
        string declaration = source.Substring(declarationStart, declarationEnd - declarationStart);
        int encodingStart = declaration.IndexOf("encoding", StringComparison.OrdinalIgnoreCase);
        if (encodingStart < 0) return fallback;
        int equals = declaration.IndexOf('=', encodingStart + 8);
        if (equals < 0) return fallback;
        int quote = equals + 1;
        while (quote < declaration.Length && char.IsWhiteSpace(declaration[quote])) quote++;
        if (quote >= declaration.Length || declaration[quote] != '\'' && declaration[quote] != '"') return fallback;
        char delimiter = declaration[quote++];
        int valueEnd = declaration.IndexOf(delimiter, quote);
        if (valueEnd <= quote) return fallback;
        try {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(declaration.Substring(quote, valueEnd - quote), EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        } catch (ArgumentException) {
            return fallback;
        }
    }

    internal static byte[] RemovePreamble(byte[] bytes, Encoding encoding) {
        byte[] preamble = encoding.GetPreamble();
        if (preamble.Length == 0 || bytes.Length < preamble.Length) return bytes;
        for (int index = 0; index < preamble.Length; index++) if (bytes[index] != preamble[index]) return bytes;
        var result = new byte[bytes.Length - preamble.Length];
        Buffer.BlockCopy(bytes, preamble.Length, result, 0, result.Length);
        return result;
    }
}
