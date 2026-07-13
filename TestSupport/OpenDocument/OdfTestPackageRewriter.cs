using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace OfficeIMO.OpenDocument.Testing;

/// <summary>Rebuilds deliberately modified test packages while preserving the required ODF ZIP envelope.</summary>
internal static class OdfTestPackageRewriter {
    internal static byte[] Rewrite(byte[] sourceBytes, IReadOnlyList<OdfTestPackageEntry>? replacements = null,
        IReadOnlyList<OdfTestPackageEntry>? additions = null) {
        var replacementMap = (replacements ?? Array.Empty<OdfTestPackageEntry>())
            .ToDictionary(item => item.Name, StringComparer.Ordinal);
        return RewriteCore(sourceBytes,
            (name, bytes) => replacementMap.TryGetValue(name, out OdfTestPackageEntry? replacement) ? replacement.Bytes : bytes,
            static _ => true,
            additions);
    }

    internal static byte[] Rewrite(byte[] sourceBytes, Func<string, byte[], byte[]> transform) {
        if (transform == null) throw new ArgumentNullException(nameof(transform));
        return RewriteCore(sourceBytes, transform, static _ => true, additions: null);
    }

    internal static byte[] Remove(byte[] sourceBytes, string removedPath) {
        if (removedPath == null) throw new ArgumentNullException(nameof(removedPath));
        return RewriteCore(sourceBytes, static (_, bytes) => bytes,
            name => !string.Equals(name, removedPath, StringComparison.Ordinal), additions: null);
    }

    private static byte[] RewriteCore(byte[] sourceBytes, Func<string, byte[], byte[]> transform,
        Func<string, bool> include, IReadOnlyList<OdfTestPackageEntry>? additions) {
        if (sourceBytes == null) throw new ArgumentNullException(nameof(sourceBytes));

        var entries = new List<OdfZipWriteEntry>();
        using (var sourceStream = new MemoryStream(sourceBytes, writable: false))
        using (var source = new ZipArchive(sourceStream, ZipArchiveMode.Read)) {
            foreach (ZipArchiveEntry entry in source.Entries) {
                if (!include(entry.FullName)) continue;
                byte[] bytes = transform(entry.FullName, ReadEntry(entry));
                entries.Add(CreateWriteEntry(entry.FullName, bytes));
            }
        }
        foreach (OdfTestPackageEntry addition in additions ?? Array.Empty<OdfTestPackageEntry>()) {
            entries.Add(CreateWriteEntry(addition.Name, addition.Bytes));
        }
        return OdfZipWriter.Write(entries, deterministic: true);
    }

    private static OdfZipWriteEntry CreateWriteEntry(string name, byte[] bytes) {
        bool compress = name != "mimetype" && !name.EndsWith("/", StringComparison.Ordinal);
        return new OdfZipWriteEntry(name, bytes, compress);
    }

    private static byte[] ReadEntry(ZipArchiveEntry entry) {
        using Stream input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }
}

internal sealed class OdfTestPackageEntry {
    internal OdfTestPackageEntry(string name, byte[] bytes) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
    }

    internal string Name { get; }
    internal byte[] Bytes { get; }
}
