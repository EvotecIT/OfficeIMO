using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static class EmailByteReader {
    internal static byte[] ReadAll(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        return ReadAll(stream, maximumBytes, nameof(EmailReaderOptions.MaxInputBytes), cancellationToken);
    }

    internal static byte[] ReadAll(Stream stream, long maximumBytes, string limitName,
        CancellationToken cancellationToken) {
        try {
            return OfficeStreamReader.ReadAllBytes(stream, cancellationToken, maximumBytes);
        } catch (InvalidDataException) {
            throw CreateInputLimitException(stream, maximumBytes, limitName);
        }
    }

    internal static async Task<byte[]> ReadAllAsync(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        return await ReadAllAsync(stream, maximumBytes, nameof(EmailReaderOptions.MaxInputBytes), cancellationToken)
            .ConfigureAwait(false);
    }

    internal static async Task<byte[]> ReadAllAsync(Stream stream, long maximumBytes, string limitName,
        CancellationToken cancellationToken) {
        try {
            return await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, maximumBytes).ConfigureAwait(false);
        } catch (InvalidDataException) {
            throw CreateInputLimitException(stream, maximumBytes, limitName);
        }
    }

    private static EmailLimitExceededException CreateInputLimitException(Stream stream, long maximumBytes,
        string limitName) =>
        new EmailLimitExceededException(
            limitName,
            stream.CanSeek ? stream.Length : checked(maximumBytes + 1),
            maximumBytes);
}
