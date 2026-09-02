namespace OfficeIMO.Workflows;

internal static class OfficeWorkflowInputReader {
    private const int BufferSize = 81920;

    internal static byte[] ReadAllBytes(
        string path,
        long maximumInputBytes,
        CancellationToken cancellationToken) {
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize, FileOptions.SequentialScan);
        return ReadAllBytes(source, Path.GetFileName(path), maximumInputBytes, cancellationToken);
    }

    internal static byte[] ReadAllBytes(
        Stream source,
        string sourceName,
        long maximumInputBytes,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(source);
        if (!source.CanRead) throw new ArgumentException("Stream must be readable.", nameof(source));
        if (maximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));

        cancellationToken.ThrowIfCancellationRequested();
        if (source.CanSeek && source.Length - source.Position > maximumInputBytes) {
            throw CreateLimitException(sourceName, source.Length - source.Position, maximumInputBytes);
        }

        int initialCapacity = source.CanSeek
            ? checked((int)Math.Min(Math.Min(source.Length - source.Position, maximumInputBytes), 1024L * 1024L))
            : 0;
        using var output = new MemoryStream(initialCapacity);
        var buffer = new byte[BufferSize];
        long total = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            cancellationToken.ThrowIfCancellationRequested();
            if (read == 0) break;

            total = checked(total + read);
            if (total > maximumInputBytes) {
                throw CreateLimitException(sourceName, total, maximumInputBytes);
            }
            if (total > Array.MaxLength) {
                throw new InvalidOperationException(
                    $"Input '{sourceName}' cannot be materialized as one byte array because it exceeds the runtime array limit.");
            }
            output.Write(buffer, 0, read);
        }

        return output.ToArray();
    }

    private static InvalidOperationException CreateLimitException(string sourceName, long size, long maximumInputBytes) =>
        new($"Input '{sourceName}' is at least {size:N0} bytes, above the configured {maximumInputBytes:N0}-byte limit.");
}
