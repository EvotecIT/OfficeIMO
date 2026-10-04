using OfficeIMO.Reader;

namespace OfficeIMO.Tool.Commands.Reader;

internal static class ReaderToolFileDiscovery {
    internal static IReadOnlyList<string> FindSupportedFiles(
        string rootPath,
        OfficeDocumentReader reader,
        bool recurse,
        int maxFiles,
        long? maxTotalBytes,
        CancellationToken cancellationToken) =>
        reader.EnumerateDocumentPaths(new[] { Path.GetFullPath(rootPath) },
            new ReaderFolderOptions {
                Recurse = recurse,
                MaxFiles = maxFiles,
                MaxTotalBytes = maxTotalBytes,
                SkipReparsePoints = true,
                DeterministicOrder = true
            }, cancellationToken)
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();
}
