using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Pdf;

/// <summary>Creates owner-only, non-shareable spill storage that the OS deletes on close.</summary>
internal static class PdfTemporaryFile {
    internal static FileStream Create(string suffix, FileOptions options, out string path) =>
        OfficeTemporaryFile.Create("OfficeIMO.Pdf-", suffix, options, out path);
}
