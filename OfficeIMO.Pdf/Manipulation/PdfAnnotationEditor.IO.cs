using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath = Path.GetFullPath(outputPath);
        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        return fullPath;
    }

    private static void WriteFile(string outputPath, byte[] bytes, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        OfficeFileCommit.Write(outputPath, stream => OfficeStreamWriter.CopyBytes(stream, bytes, cancellationToken), cancellationToken);
    }
}
