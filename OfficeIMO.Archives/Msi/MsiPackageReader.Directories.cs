using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    private MsiResolvedDirectory ResolveDirectory(string id, Dictionary<string, object?[]> directories,
        Dictionary<string, MsiResolvedDirectory> cache) {
        if (cache.TryGetValue(id, out MsiResolvedDirectory cached)) return cached;
        var chain = new List<KeyValuePair<string, object?[]>>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        string current = id;
        var resolved = new MsiResolvedDirectory(string.Empty, 0, 0);
        while (true) {
            _token.ThrowIfCancellationRequested();
            if (cache.TryGetValue(current, out resolved)) break;
            if (!visited.Add(current) || chain.Count >= _limits.Depth)
                throw new InvalidDataException("MSI directory graph cycles or exceeds configured structural depth.");
            if (!directories.TryGetValue(current, out object?[]? row)) throw new InvalidDataException("MSI references an absent directory.");
            chain.Add(new KeyValuePair<string, object?[]>(current, row));
            string? parent = row[1] as string;
            if (parent == null || string.Equals(parent, current, StringComparison.Ordinal)) {
                resolved = new MsiResolvedDirectory(string.Empty, 0, 0);
                break;
            }
            current = parent;
        }
        // The cached prefix carries its complete structural depth, including
        // directories authored as '.', so cache hits cannot reset the budget.
        if (resolved.StructuralDepth > _limits.Depth - chain.Count)
            throw new InvalidDataException("MSI directory graph exceeds configured structural depth.");
        for (int index = chain.Count - 1; index >= 0; index--) {
            _token.ThrowIfCancellationRequested();
            var entry = chain[index];
            string definition = RequiredString(entry.Value[2], "Directory.DefaultDir");
            string[] pair = definition.Split(':');
            if (pair.Length > 2) throw new InvalidDataException("MSI directory contains an invalid target/source name pair.");
            string target = LongName(pair[0]);
            string? parent = entry.Value[1] as string;
            bool root = parent == null || string.Equals(parent, entry.Key, StringComparison.Ordinal);
            int structuralDepth = resolved.StructuralDepth + 1;
            if (!root && target != ".") {
                ArchivePaths.ValidateSegment(target);
                int segments = resolved.PathSegments + 1;
                long length = (long)resolved.Path.Length + (resolved.Path.Length == 0 ? 0 : 1) + target.Length;
                if (segments > _limits.Depth || length > _limits.PathLength)
                    throw new InvalidDataException("MSI authored directory path exceeds configured bounds.");
                // Check all bounds before allocating or caching the assembled prefix.
                string path = resolved.Path.Length == 0 ? target : resolved.Path + "/" + target;
                resolved = new MsiResolvedDirectory(path, structuralDepth, segments);
            } else resolved = new MsiResolvedDirectory(resolved.Path, structuralDepth, resolved.PathSegments);
            cache.Add(entry.Key, resolved);
        }
        return resolved;
    }

    private readonly struct MsiResolvedDirectory {
        internal MsiResolvedDirectory(string path, int structuralDepth, int pathSegments) {
            Path = path; StructuralDepth = structuralDepth; PathSegments = pathSegments;
        }
        internal string Path { get; }
        internal int StructuralDepth { get; }
        internal int PathSegments { get; }
    }
}
