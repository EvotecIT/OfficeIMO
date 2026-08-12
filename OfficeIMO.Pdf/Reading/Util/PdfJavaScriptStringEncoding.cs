namespace OfficeIMO.Pdf;

/// <summary>Text-string encoding used by named PDF JavaScript keys and sources.</summary>
internal static class PdfJavaScriptStringEncoding {
    internal static bool TryDecode(byte[] bytes, out string value) {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
            try {
                value = new System.Text.UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3);
                return IsWellFormedUtf16(value);
            } catch (System.Text.DecoderFallbackException) {
                value = string.Empty;
                return false;
            }
        }
        if (bytes.Length >= 2 &&
            (bytes[0] == 0xFE && bytes[1] == 0xFF || bytes[0] == 0xFF && bytes[1] == 0xFE)) {
            if (((bytes.Length - 2) & 1) != 0) {
                value = string.Empty;
                return false;
            }
            try {
                value = PdfTextString.Decode(bytes);
                return IsWellFormedUtf16(value);
            } catch {
                value = string.Empty;
                return false;
            }
        }
        return PdfDocEncoding.TryDecode(bytes, out value);
    }

    internal static byte[] EncodeUnicode(string value, string parameterName) {
        if (!IsWellFormedUtf16(value)) {
            throw new ArgumentException("PDF JavaScript text must contain well-formed Unicode.", parameterName);
        }
        var bytes = new byte[2 + checked(value.Length * 2)];
        bytes[0] = 0xFE;
        bytes[1] = 0xFF;
        for (int i = 0; i < value.Length; i++) {
            bytes[2 + (i * 2)] = (byte)(value[i] >> 8);
            bytes[3 + (i * 2)] = (byte)value[i];
        }
        return bytes;
    }

    private static bool IsWellFormedUtf16(string value) {
        for (int i = 0; i < value.Length; i++) {
            char character = value[i];
            if (char.IsHighSurrogate(character)) {
                if (i + 1 >= value.Length || !char.IsLowSurrogate(value[++i])) return false;
            } else if (char.IsLowSurrogate(character)) {
                return false;
            }
        }
        return true;
    }
}
