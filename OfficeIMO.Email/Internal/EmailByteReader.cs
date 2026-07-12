namespace OfficeIMO.Email;

internal static class EmailByteReader {
    internal static byte[] ReadAll(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The stream must be readable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();

        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            EnsureWithinLimit(remaining, maximumBytes);
        }

        using (MemoryStream output = new MemoryStream()) {
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                total += read;
                EnsureWithinLimit(total, maximumBytes);
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }
    }

    internal static async Task<byte[]> ReadAllAsync(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The stream must be readable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();

        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            EnsureWithinLimit(remaining, maximumBytes);
        }

        using (MemoryStream output = new MemoryStream()) {
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) != 0) {
                total += read;
                EnsureWithinLimit(total, maximumBytes);
                await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
            }
            return output.ToArray();
        }
    }

    private static void EnsureWithinLimit(long value, long maximumBytes) {
        if (value > maximumBytes) throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), value, maximumBytes);
    }
}
