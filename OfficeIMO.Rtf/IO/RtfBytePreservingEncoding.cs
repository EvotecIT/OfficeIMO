using OfficeIMO.Drawing.Internal;
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
        return FromBytes(ReadBytesToEnd(stream, null, CancellationToken.None));
    }

    public static async Task<string> ReadToEndAsync(Stream stream, CancellationToken cancellationToken) {
        return FromBytes(await ReadBytesToEndAsync(stream, null, cancellationToken).ConfigureAwait(false));
    }

    public static byte[] ReadBytesToEnd(Stream stream, long? maxBytes, CancellationToken cancellationToken) {
        try {
            return OfficeStreamReader.ReadAllBytes(stream, cancellationToken, maxBytes);
        } catch (InvalidDataException) when (maxBytes.HasValue) {
            throw CreateInputLimitException(stream, maxBytes.Value);
        }
    }

    public static async Task<byte[]> ReadBytesToEndAsync(Stream stream, long? maxBytes, CancellationToken cancellationToken) {
        try {
            return await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, maxBytes).ConfigureAwait(false);
        } catch (InvalidDataException) when (maxBytes.HasValue) {
            throw CreateInputLimitException(stream, maxBytes.Value);
        }
    }

    public static void WriteAllText(string path, string rtf) {
        OfficeFileCommit.WriteAllBytes(path, ToBytes(rtf));
    }

    public static async Task WriteAllTextAsync(string path, string rtf, CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        byte[] bytes = ToBytes(rtf);
        await OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    public static void WriteTo(Stream stream, string rtf) {
        byte[] bytes = ToBytes(rtf);
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
    }

    public static async Task WriteToAsync(Stream stream, string rtf, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        byte[] bytes = ToBytes(rtf);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    private static string FromBytes(byte[] bytes) {
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) {
            chars[i] = (char)bytes[i];
        }

        return new string(chars);
    }

    private static RtfReadLimitException CreateInputLimitException(Stream stream, long limit) {
        long actual = stream.CanSeek ? stream.Length : checked(limit + 1);
        return new RtfReadLimitException(
            "RtfInputByteLimitExceeded",
            $"RTF input exceeded {nameof(RtfReadOptions.MaxInputBytes)} ({actual} > {limit}).",
            nameof(RtfReadOptions.MaxInputBytes),
            actual,
            limit);
    }

    internal static byte[] ToBytes(string rtf) {
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
