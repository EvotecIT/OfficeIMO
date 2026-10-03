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
        long totalBytes = 0L;
        foreach (string path in paths) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Paths cannot contain null or empty values.", nameof(paths));
            }

            if (!Directory.Exists(path)) {
                yield return path;
                continue;
            }

            if (IsRegisteredDirectoryBundle(path)) {
                if (effectiveFolder.MaxTotalBytes.HasValue) {
                    if (!TryGetKnownDocumentLength(path, effectiveFolder.MaxTotalBytes.Value - totalBytes,
                            cancellationToken, out long length)) continue;
                    totalBytes += length;
                }
                yield return NormalizeDirectoryPackagePath(path);
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

                if (effectiveFolder.MaxTotalBytes.HasValue) {
                    if (!TryGetKnownDocumentLength(file, effectiveFolder.MaxTotalBytes.Value - totalBytes,
                            cancellationToken, out long fileLength) ||
                        fileLength > effectiveFolder.MaxTotalBytes.Value - totalBytes) {
                        continue;
                    }

                    totalBytes += fileLength;
                }

                filesEnumerated++;
                yield return file;
            }
        }
    }

    private static bool TryGetKnownDocumentLength(string path, long maximumBytes,
        CancellationToken cancellationToken, out long length) {
        if (Directory.Exists(path) && IsRegisteredDirectoryBundle(path)) {
            length = 0;
            if (maximumBytes <= 0) return false;
            try {
                OfficeDocumentReadResult result = ReadDirectoryBundle(path,
                    ApplyFolderInputBudget(new ReaderOptions { ComputeHashes = false }, maximumBytes), cancellationToken);
                length = result.Source.LengthBytes!.Value;
                return true;
            } catch (OperationCanceledException) {
                throw;
            } catch (InvalidDataException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (NotSupportedException) {
                return false;
            }
        }
        return TryGetKnownFileLength(path, out length);
    }

    private static bool TryGetKnownFileLength(string path, out long length) {
        try {
            length = new FileInfo(path).Length;
            return length >= 0L;
        } catch {
            length = 0L;
            return false;
        }
    }
}
