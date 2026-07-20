using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static IEnumerable<string> EnumerateDocumentPaths(
        IEnumerable<string> paths,
        ReaderFolderOptions? folderOptions,
        CancellationToken cancellationToken) {
        if (paths == null) throw new ArgumentNullException(nameof(paths));

        ReaderFolderOptions effectiveFolder = NormalizeFolderOptions(folderOptions);
        HashSet<string> allowedExtensions = NormalizeExtensions(effectiveFolder.Extensions);
        foreach (string path in paths) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Paths cannot contain null or empty values.", nameof(paths));
            }

            if (!Directory.Exists(path)) {
                yield return path;
                continue;
            }

            int filesEnumerated = 0;
            foreach (string file in EnumerateFilesSafeDeterministic(path, effectiveFolder, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (filesEnumerated >= effectiveFolder.MaxFiles) break;

                string extension = NormalizeExtension(TryGetExtension(file));
                if (!allowedExtensions.Contains(extension) &&
                    !string.Equals(Path.GetFileName(file), "winmail.dat", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                filesEnumerated++;
                yield return file;
            }
        }
    }
}
