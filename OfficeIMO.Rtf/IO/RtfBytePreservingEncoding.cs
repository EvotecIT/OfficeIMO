namespace OfficeIMO.Rtf;

internal static class RtfBytePreservingEncoding {
    public static string ReadAllText(string path) {
        byte[] bytes = File.ReadAllBytes(path);
        return FromBytes(bytes);
    }

    public static string GetString(byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return FromBytes(bytes);
    }

    public static string ReadToEnd(Stream stream) {
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return FromBytes(memory.ToArray());
    }

    public static byte[] GetBytes(string rtf) => ToBytes(rtf);

    public static void WriteAllText(string path, string rtf) {
        File.WriteAllBytes(path, ToBytes(rtf));
    }

    public static void WriteTo(Stream stream, string rtf) {
        byte[] bytes = ToBytes(rtf);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string FromBytes(byte[] bytes) {
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) {
            chars[i] = (char)bytes[i];
        }

        return new string(chars);
    }

    private static byte[] ToBytes(string rtf) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));

        var bytes = new byte[rtf.Length];
        for (int i = 0; i < rtf.Length; i++) {
            char value = rtf[i];
            if (value > byte.MaxValue) {
                throw new InvalidOperationException("Lossless RTF byte output can only write source-preserved characters in the 0-255 byte range.");
            }

            bytes[i] = (byte)value;
        }

        return bytes;
    }
}
