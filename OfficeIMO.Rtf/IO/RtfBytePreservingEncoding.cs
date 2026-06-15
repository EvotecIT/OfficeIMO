namespace OfficeIMO.Rtf;

internal static class RtfBytePreservingEncoding {
    public static string ReadAllText(string path) {
        byte[] bytes = File.ReadAllBytes(path);
        return FromBytes(bytes);
    }

    public static async Task<string> ReadAllTextAsync(string path, CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true);
        return await ReadToEndAsync(stream, cancellationToken).ConfigureAwait(false);
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

    public static async Task<string> ReadToEndAsync(Stream stream, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        using var memory = new MemoryStream();
        await stream.CopyToAsync(memory, 81920, cancellationToken).ConfigureAwait(false);
        return FromBytes(memory.ToArray());
    }

    public static byte[] GetBytes(string rtf) => ToBytes(rtf);

    public static void WriteAllText(string path, string rtf) {
        File.WriteAllBytes(path, ToBytes(rtf));
    }

    public static async Task WriteAllTextAsync(string path, string rtf, CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        byte[] bytes = ToBytes(rtf);
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    public static void WriteTo(Stream stream, string rtf) {
        byte[] bytes = ToBytes(rtf);
        stream.Write(bytes, 0, bytes.Length);
    }

    public static async Task WriteToAsync(Stream stream, string rtf, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        byte[] bytes = ToBytes(rtf);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
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
