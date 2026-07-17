using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Builds deterministic CFB directory entries from slash-separated stream paths.
    /// </summary>
    internal sealed class OfficeCompoundWriterLayout {
        private const uint NoStream = 0xffffffff;
        private readonly List<OfficeCompoundWriterEntry> _entries = new List<OfficeCompoundWriterEntry>();

        private OfficeCompoundWriterLayout(OfficeCompoundWriterEntry root) {
            Root = root;
        }

        internal OfficeCompoundWriterEntry Root { get; }

        internal IReadOnlyList<OfficeCompoundWriterEntry> Entries => _entries;

        internal IReadOnlyList<OfficeCompoundWriterEntry> Streams => _entries.Where(entry => entry.ObjectType == 2).ToArray();

        internal static OfficeCompoundWriterLayout Create(IReadOnlyList<OfficeCompoundStream> streams,
            OfficeCompoundFile? source = null,
            IReadOnlyCollection<string>? removedPaths = null) {
            var root = new OfficeCompoundWriterEntry("Root Entry", string.Empty, 5, null);
            var layout = new OfficeCompoundWriterLayout(root);
            if (source != null) {
                root.ApplyMetadata(source.RootEntry);
                foreach (OfficeCompoundFileEntry storage in source.Entries.Where(entry =>
                             entry.ObjectType == 1 && !entry.IsFallback
                             && !IsRemoved(entry.Path, removedPaths))
                         .OrderBy(entry => entry.Path.Count(character => character == '/'))
                         .ThenBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase)) {
                    layout.AddStorage(storage.Path);
                }
            }
            foreach (OfficeCompoundStream stream in streams) layout.AddStream(stream);
            if (source != null) layout.ApplyMetadata(source);
            layout.AssignDirectoryEntries();
            layout.AssignTreeLinks(root);
            return layout;
        }

        private static bool IsRemoved(string path, IReadOnlyCollection<string>? removedPaths) {
            if (removedPaths == null) return false;
            foreach (string removedPath in removedPaths) {
                if (string.Equals(path, removedPath, StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith(removedPath + "/", StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }
            return false;
        }

        private void AddStorage(string storagePath) {
            string normalized = storagePath.Replace('\\', '/').Trim('/');
            if (normalized.Length == 0) return;
            string[] segments = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            OfficeCompoundWriterEntry parent = Root;
            string path = string.Empty;
            foreach (string name in segments) {
                ValidateName(name);
                path = path.Length == 0 ? name : string.Concat(path, "/", name);
                OfficeCompoundWriterEntry? existing = parent.Children.FirstOrDefault(child =>
                    string.Equals(child.Name, name, StringComparison.OrdinalIgnoreCase));
                if (existing != null) {
                    if (existing.ObjectType != 1) {
                        throw new ArgumentException(string.Concat("Conflicting compound storage path '", path, "'."),
                            nameof(storagePath));
                    }
                    parent = existing;
                    continue;
                }
                var created = new OfficeCompoundWriterEntry(name, path, 1, null);
                parent.Children.Add(created);
                parent = created;
            }
        }

        private void ApplyMetadata(OfficeCompoundFile source) {
            var metadata = source.Entries.Where(entry => !entry.IsFallback)
                .GroupBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            ApplyMetadata(Root, metadata);
        }

        private static void ApplyMetadata(OfficeCompoundWriterEntry entry,
            IReadOnlyDictionary<string, OfficeCompoundFileEntry> metadata) {
            if (entry.Path.Length > 0 && metadata.TryGetValue(entry.Path, out OfficeCompoundFileEntry? sourceEntry)) {
                entry.ApplyMetadata(sourceEntry);
            }
            foreach (OfficeCompoundWriterEntry child in entry.Children) ApplyMetadata(child, metadata);
        }

        private void AddStream(OfficeCompoundStream stream) {
            string normalized = stream.Name.Replace('\\', '/').Trim('/');
            if (normalized.Length == 0) throw new ArgumentException("Compound stream path is required.", nameof(stream));
            string[] segments = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            OfficeCompoundWriterEntry parent = Root;
            string path = string.Empty;
            for (int i = 0; i < segments.Length; i++) {
                string name = segments[i];
                ValidateName(name);
                path = path.Length == 0 ? name : string.Concat(path, "/", name);
                bool isStream = i + 1 == segments.Length;
                OfficeCompoundWriterEntry? existing = parent.Children.FirstOrDefault(child =>
                    string.Equals(child.Name, name, StringComparison.OrdinalIgnoreCase));
                if (existing != null) {
                    if (isStream || existing.ObjectType != 1) {
                        throw new ArgumentException(string.Concat("Duplicate or conflicting compound path '", path, "'."), nameof(stream));
                    }
                    parent = existing;
                    continue;
                }

                var created = new OfficeCompoundWriterEntry(name, path, isStream ? (byte)2 : (byte)1,
                    isStream ? stream : (OfficeCompoundStream?)null);
                parent.Children.Add(created);
                parent = created;
            }
        }

        private void AssignDirectoryEntries() {
            _entries.Clear();
            AssignDirectoryEntry(Root);
        }

        private void AssignDirectoryEntry(OfficeCompoundWriterEntry entry) {
            entry.DirectoryIndex = _entries.Count;
            _entries.Add(entry);
            foreach (OfficeCompoundWriterEntry child in entry.Children.OrderBy(item => item.Name, OfficeCompoundDirectoryNameComparer.Instance)) {
                AssignDirectoryEntry(child);
            }
        }

        private void AssignTreeLinks(OfficeCompoundWriterEntry storage) {
            OfficeCompoundWriterEntry[] children = storage.Children
                .OrderBy(entry => entry.Name, OfficeCompoundDirectoryNameComparer.Instance)
                .ToArray();
            storage.ChildId = BuildBalancedTree(children, 0, children.Length - 1);
            foreach (OfficeCompoundWriterEntry child in children.Where(entry => entry.ObjectType == 1)) AssignTreeLinks(child);
        }

        private static uint BuildBalancedTree(OfficeCompoundWriterEntry[] entries, int first, int last) {
            if (first > last) return NoStream;
            int middle = first + ((last - first) / 2);
            OfficeCompoundWriterEntry entry = entries[middle];
            entry.LeftSiblingId = BuildBalancedTree(entries, first, middle - 1);
            entry.RightSiblingId = BuildBalancedTree(entries, middle + 1, last);
            return unchecked((uint)entry.DirectoryIndex);
        }

        private static void ValidateName(string name) {
            if (name.Length == 0) throw new ArgumentException("Compound entry name is required.", nameof(name));
            if (name.Length > 31) throw new ArgumentException("Compound entry names cannot exceed 31 UTF-16 characters.", nameof(name));
            if (name.IndexOfAny(new[] { '/', '\\', ':', '!' }) >= 0) {
                throw new ArgumentException(string.Concat("Compound entry name '", name, "' contains a reserved character."), nameof(name));
            }
        }
    }

    internal sealed class OfficeCompoundWriterEntry {
        internal OfficeCompoundWriterEntry(string name, string path, byte objectType, OfficeCompoundStream? stream) {
            Name = name;
            Path = path;
            ObjectType = objectType;
            Stream = stream;
        }

        internal string Name { get; }

        internal string Path { get; }

        internal byte ObjectType { get; }

        internal OfficeCompoundStream? Stream { get; }

        internal List<OfficeCompoundWriterEntry> Children { get; } = new List<OfficeCompoundWriterEntry>();

        internal Guid ClassId { get; private set; }

        internal uint StateBits { get; private set; }

        internal ulong CreationTime { get; private set; }

        internal ulong ModifiedTime { get; private set; }

        internal void ApplyMetadata(OfficeCompoundFileEntry entry) {
            ClassId = entry.ClassId;
            StateBits = entry.StateBits;
            CreationTime = entry.CreationTime;
            ModifiedTime = entry.ModifiedTime;
        }

        internal int DirectoryIndex { get; set; }

        internal uint LeftSiblingId { get; set; } = 0xffffffff;

        internal uint RightSiblingId { get; set; } = 0xffffffff;

        internal uint ChildId { get; set; } = 0xffffffff;
    }

    internal sealed class OfficeCompoundDirectoryNameComparer : IComparer<string> {
        internal static OfficeCompoundDirectoryNameComparer Instance { get; } = new OfficeCompoundDirectoryNameComparer();

        public int Compare(string? left, string? right) {
            int length = (left?.Length ?? 0).CompareTo(right?.Length ?? 0);
            if (length != 0) return length;
            int ignoreCase = StringComparer.OrdinalIgnoreCase.Compare(left, right);
            return ignoreCase != 0 ? ignoreCase : StringComparer.Ordinal.Compare(left, right);
        }
    }
}
