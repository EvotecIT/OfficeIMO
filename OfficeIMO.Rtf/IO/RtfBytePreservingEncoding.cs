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
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        using var memory = CreateInputBuffer(stream, maxBytes);
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total += read;
            EnforceInputByteLimit(total, maxBytes);
            memory.Write(buffer, 0, read);
        }

        return memory.ToArray();
    }

    public static async Task<byte[]> ReadBytesToEndAsync(Stream stream, long? maxBytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        using var memory = CreateInputBuffer(stream, maxBytes);
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            total += read;
            EnforceInputByteLimit(total, maxBytes);
            await memory.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }

        return memory.ToArray();
    }

    public static byte[] GetBytes(string rtf) => ToBytes(rtf);

    public static void WriteAllText(string path, string rtf) {
        OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(path, ToBytes(rtf));
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

    private static MemoryStream CreateInputBuffer(Stream stream, long? maxBytes) {
        if (!stream.CanSeek) return new MemoryStream();
        long remaining = Math.Max(0, stream.Length - stream.Position);
        EnforceInputByteLimit(remaining, maxBytes);
        return remaining <= int.MaxValue ? new MemoryStream((int)remaining) : new MemoryStream();
    }

    private static void EnforceInputByteLimit(long actual, long? limit) {
        if (limit.HasValue && actual > limit.Value) {
            throw new RtfReadLimitException(
                "RtfInputByteLimitExceeded",
                $"RTF input exceeded {nameof(RtfReadOptions.MaxInputBytes)} ({actual} > {limit.Value}).",
                nameof(RtfReadOptions.MaxInputBytes),
                actual,
                limit.Value);
        }
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
