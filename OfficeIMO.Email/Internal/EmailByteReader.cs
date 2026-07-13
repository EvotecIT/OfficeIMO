using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal static class EmailByteReader {
    internal static byte[] ReadAll(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        try {
            return OfficeStreamReader.ReadAllBytes(stream, cancellationToken, maximumBytes);
        } catch (InvalidDataException) {
            throw CreateInputLimitException(stream, maximumBytes);
        }
    }

    internal static async Task<byte[]> ReadAllAsync(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        try {
            return await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, maximumBytes).ConfigureAwait(false);
        } catch (InvalidDataException) {
            throw CreateInputLimitException(stream, maximumBytes);
        }
    }

    private static EmailLimitExceededException CreateInputLimitException(Stream stream, long maximumBytes) =>
        new EmailLimitExceededException(
            nameof(EmailReaderOptions.MaxInputBytes),
            stream.CanSeek ? stream.Length : checked(maximumBytes + 1),
            maximumBytes);
}
