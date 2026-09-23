namespace OfficeIMO.Pdf;

/// <summary>Text-string encoding used by named PDF JavaScript keys and sources.</summary>
internal static class PdfJavaScriptStringEncoding {
    internal static bool TryDecode(byte[] bytes, out string value, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
            try {
                var decoder = new System.Text.UTF8Encoding(false, true).GetDecoder();
                var characters = new char[8194];
                var decoded = new System.Text.StringBuilder();
                for (int offset = 3; offset < bytes.Length;) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int count = Math.Min(8192, bytes.Length - offset);
                    int written = decoder.GetChars(bytes, offset, count, characters, 0, offset + count == bytes.Length);
                    decoded.Append(characters, 0, written);
                    offset += count;
                }
                cancellationToken.ThrowIfCancellationRequested();
                value = decoded.ToString();
                return IsWellFormedUtf16(value, cancellationToken);
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
            var characters = new char[(bytes.Length - 2) / 2];
            bool bigEndian = bytes[0] == 0xFE;
            for (int index = 0; index < characters.Length; index++) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                int offset = 2 + index * 2;
                characters[index] = bigEndian
                    ? (char)((bytes[offset] << 8) | bytes[offset + 1])
                    : (char)(bytes[offset] | (bytes[offset + 1] << 8));
            }
            cancellationToken.ThrowIfCancellationRequested();
            value = new string(characters);
            return IsWellFormedUtf16(value, cancellationToken);
        }
        return PdfDocEncoding.TryDecode(bytes, out value, cancellationToken);
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

    private static bool IsWellFormedUtf16(string value, System.Threading.CancellationToken cancellationToken = default) {
        for (int i = 0; i < value.Length; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
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
