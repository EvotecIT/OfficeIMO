using OfficeIMO.Drawing.Internal;
using System;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Drawing;

internal static class OfficeImageExportPath {
    internal static string ResolveFile(
        string path,
        OfficeImageExportFormat format,
        OfficeImageExportFileConflictPolicy conflictPolicy) {
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

        if (conflictPolicy != OfficeImageExportFileConflictPolicy.CreateUnique || !File.Exists(resolved)) return resolved;
        return CreateUniquePath(resolved);
    }

    internal static OfficeFileCommit.ConflictPolicy ToCommitPolicy(OfficeImageExportFileConflictPolicy policy) =>
        policy == OfficeImageExportFileConflictPolicy.Replace
            ? OfficeFileCommit.ConflictPolicy.Replace
            : OfficeFileCommit.ConflictPolicy.FailIfExists;

    internal static string CreateUniquePath(string path) {
        if (!File.Exists(path)) return path;
        string? directory = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
        string baseName = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);
        for (int suffix = 2; suffix < int.MaxValue; suffix++) {
            string candidate = Path.Combine(directory, baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension);
            if (!File.Exists(candidate)) return candidate;
        }

        throw new IOException("Could not allocate a unique image export destination.");
    }
}
