using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

/// <summary>Bounds each caller-supplied font face before it is buffered or copied.</summary>
internal static class PdfFontInput {
    internal const long MaximumFontFaceBytes = 128L * 1024L * 1024L;

    internal static byte[] ReadFile(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        string fullPath = Path.GetFullPath(path);
        // The opened file's length comes from the descriptor. FileInfo.Length would read the entry's attributes,
        // which fails on iOS for files the app doesn't own, such as system fonts.
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        EnsureWithinLimit(stream.Length);
        return OfficeStreamReader.ReadRemainingBytes(stream, MaximumFontFaceBytes);
    }

    internal static void EnsureWithinLimit(long length) {
        if (length > MaximumFontFaceBytes) {
            throw new InvalidDataException($"Font input is {length} bytes and exceeds the {MaximumFontFaceBytes} byte limit.");
        }
    }
}
