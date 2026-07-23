namespace OfficeIMO.Pdf;

/// <summary>Creates owner-only, non-shareable spill storage that the OS deletes on close.</summary>
internal static class PdfTemporaryFile {
    internal static FileStream Create(string suffix, FileOptions options, out string path) {
        path = Path.Combine(Path.GetTempPath(),
            "OfficeIMO.Pdf-" + Guid.NewGuid().ToString("N") + suffix);
#if NET6_0_OR_GREATER
        var streamOptions = new FileStreamOptions {
            Mode = FileMode.CreateNew,
            Access = FileAccess.ReadWrite,
            Share = FileShare.None,
            BufferSize = 81920,
            Options = options | FileOptions.DeleteOnClose,
        };
        if (!OperatingSystem.IsWindows()) {
            streamOptions.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
        }
        return new FileStream(path, streamOptions);
#else
        return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.None, 81920, options | FileOptions.DeleteOnClose);
#endif
    }
}
