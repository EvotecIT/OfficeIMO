using OfficeIMO.Drawing.Internal;
using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

internal static class OfficeImageExportPath {
    internal static string NormalizeFile(string path, OfficeImageExportFormat format) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));

        string expectedExtension = format.GetFileExtension();
        string resolved = Path.GetFullPath(path);
        string extension = Path.GetExtension(resolved);
        if (string.IsNullOrEmpty(extension)) {
            resolved += expectedExtension;
        } else if (!format.HasFileExtension(extension)) {
            throw new ArgumentException(
                "Output path extension '" + extension + "' does not match the selected " + format +
                " format. Use " + expectedExtension + " or omit the extension.",
                nameof(path));
        }

        return resolved;
    }

    internal static string WriteAllBytes(
        string path,
        OfficeImageExportFormat format,
        byte[] bytes,
        OfficeImageExportFileConflictPolicy conflictPolicy) {
        string resolved = NormalizeFile(path, format);
        if (conflictPolicy != OfficeImageExportFileConflictPolicy.CreateUnique) {
            OfficeFileCommit.WriteAllBytes(resolved, bytes, ToCommitPolicy(conflictPolicy));
            return resolved;
        }

        for (int suffix = 1; suffix < int.MaxValue; suffix++) {
            string candidate = suffix == 1 ? resolved : CreateSuffixedPath(resolved, suffix);
            if (OfficeFileCommit.TryWriteAllBytes(candidate, bytes)) return candidate;
        }

        throw new IOException("Could not allocate a unique image export destination.");
    }

    internal static async Task<string> WriteAllBytesAsync(
        string path,
        OfficeImageExportFormat format,
        byte[] bytes,
        OfficeImageExportFileConflictPolicy conflictPolicy,
        CancellationToken cancellationToken) {
        string resolved = NormalizeFile(path, format);
        if (conflictPolicy != OfficeImageExportFileConflictPolicy.CreateUnique) {
            await OfficeFileCommit.WriteAllBytesAsync(
                resolved,
                bytes,
                ToCommitPolicy(conflictPolicy),
                cancellationToken).ConfigureAwait(false);
            return resolved;
        }

        for (int suffix = 1; suffix < int.MaxValue; suffix++) {
            cancellationToken.ThrowIfCancellationRequested();
            string candidate = suffix == 1 ? resolved : CreateSuffixedPath(resolved, suffix);
            if (await OfficeFileCommit.TryWriteAllBytesAsync(candidate, bytes, cancellationToken).ConfigureAwait(false)) {
                return candidate;
            }
        }

        throw new IOException("Could not allocate a unique image export destination.");
    }

    internal static OfficeFileCommit.ConflictPolicy ToCommitPolicy(OfficeImageExportFileConflictPolicy policy) =>
        policy == OfficeImageExportFileConflictPolicy.Replace
            ? OfficeFileCommit.ConflictPolicy.Replace
            : OfficeFileCommit.ConflictPolicy.FailIfExists;

    private static string CreateSuffixedPath(string path, int suffix) {
        string? directory = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
        string baseName = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);
        return Path.Combine(directory, baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension);
    }
}
