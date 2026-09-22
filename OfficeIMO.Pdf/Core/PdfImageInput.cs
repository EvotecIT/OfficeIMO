using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

/// <summary>Bounds encoded image input before PDF authoring allocates or decodes it.</summary>
internal static class PdfImageInput {
    internal const long DefaultMaximumEncodedBytes = 128L * 1024L * 1024L;

    internal static byte[] ReadFile(string path, long maximumBytes, CancellationToken cancellationToken = default) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        ValidateMaximum(maximumBytes);
        cancellationToken.ThrowIfCancellationRequested();
        string fullPath = Path.GetFullPath(path);
        var file = new FileInfo(fullPath);
        EnsureWithinLimit(file.Length, maximumBytes);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return ReadRemainingStream(stream, maximumBytes, cancellationToken);
    }

    internal static byte[] ReadRemainingStream(Stream stream, long maximumBytes, CancellationToken cancellationToken = default) {
        Guard.NotNull(stream, nameof(stream));
        ValidateMaximum(maximumBytes);
        return OfficeStreamReader.ReadRemainingBytes(stream, cancellationToken, maximumBytes);
    }

    internal static void EnsureWithinLimit(long length, long maximumBytes) {
        ValidateMaximum(maximumBytes);
        if (length > maximumBytes) {
            throw new InvalidDataException($"Encoded image input is {length} bytes and exceeds the {maximumBytes} byte limit.");
        }
    }

    internal static void ValidateMaximum(long maximumBytes) {
        if (maximumBytes < 1L || maximumBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes), "The encoded image limit must be between 1 and Int32.MaxValue bytes.");
        }
    }
}
