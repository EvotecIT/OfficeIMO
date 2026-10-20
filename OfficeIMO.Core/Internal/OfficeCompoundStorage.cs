using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Core.Internal;

/// <summary>Extracts and replaces a complete CFB storage, retaining its directory metadata.</summary>
internal static class OfficeCompoundStorage {
    internal static OfficeCompoundFile Extract(OfficeCompoundFile source, string path) {
        var root = source.Entries.SingleOrDefault(entry => entry.IsStorage && string.Equals(entry.Path, path, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException("The compound storage is missing: " + path);
        string prefix = path + "/";
        var streams = source.Streams.Where(item => item.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(item => item.Key.Substring(prefix.Length), item => item.Value, StringComparer.OrdinalIgnoreCase);
        var entries = source.Entries.Where(entry => !entry.IsFallback && entry.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(entry => Repath(entry, entry.Name, entry.Path.Substring(prefix.Length), entry.ObjectType)).ToArray();
        return new OfficeCompoundFile(streams, entries, Repath(root, "Root Entry", string.Empty, 5));
    }

    internal static OfficeCompoundFile Replace(OfficeCompoundFile source, string path, OfficeCompoundFile storage) {
        string prefix = path + "/";
        bool Inside(string value) => string.Equals(value, path, StringComparison.OrdinalIgnoreCase) || value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        var streams = source.Streams.Where(item => !Inside(item.Key)).ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);
        var entries = source.Entries.Where(entry => !Inside(entry.Path)).ToList();
        foreach (var item in storage.Streams) streams.Add(prefix + item.Key, item.Value);
        string name = path.Substring(path.LastIndexOf('/') + 1);
        entries.Add(Repath(storage.RootEntry, name, path, 1));
        foreach (var entry in storage.Entries.Where(entry => !entry.IsFallback && entry.ObjectType != 5))
            entries.Add(Repath(entry, entry.Name, prefix + entry.Path, entry.ObjectType));
        return new OfficeCompoundFile(streams, entries, source.RootEntry);
    }

    internal static OfficeCompoundFileEntry Repath(OfficeCompoundFileEntry entry, string name, string path, byte type) =>
        new(name, path, type, entry.Size, classId: entry.ClassId, stateBits: entry.StateBits,
            creationTime: entry.CreationTime, modifiedTime: entry.ModifiedTime, directoryOrder: entry.DirectoryOrder);
}
