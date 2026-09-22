using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

/// <summary>Bounds each caller-supplied font face before it is buffered or copied.</summary>
internal static class PdfFontInput {
    internal const long MaximumFontFaceBytes = 128L * 1024L * 1024L;

    internal static byte[] ReadFile(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        string fullPath = Path.GetFullPath(path);
        var file = new FileInfo(fullPath);
        EnsureWithinLimit(file.Length);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return OfficeStreamReader.ReadRemainingBytes(stream, MaximumFontFaceBytes);
    }

    internal static void EnsureWithinLimit(long length) {
        if (length > MaximumFontFaceBytes) {
            throw new InvalidDataException($"Font input is {length} bytes and exceeds the {MaximumFontFaceBytes} byte limit.");
        }
    }
}
