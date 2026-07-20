using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>
    /// Expands files and folders using this reader's registered extensions.
    /// </summary>
    /// <remarks>
    /// Explicit file paths are returned as supplied so callers receive a normal read error for missing or unsupported
    /// files. Folder contents are filtered to registered or explicitly requested extensions. Directory traversal uses
    /// <see cref="ReaderFolderOptions"/> for recursion, ordering, reparse-point handling, and per-folder limits.
    /// </remarks>
    public IEnumerable<string> EnumerateDocumentPaths(
        IEnumerable<string> paths,
        ReaderFolderOptions? folderOptions = null,
        CancellationToken cancellationToken = default) {
        return Scope(DocumentReaderEngine.EnumerateDocumentPaths(paths, folderOptions, cancellationToken));
    }
}
