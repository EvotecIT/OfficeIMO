namespace OfficeIMO.OneNote;

/// <summary>Reads the validated physical revision-store graph of desktop <c>.one</c> and <c>.onetoc2</c> files.</summary>
public static class OneNoteRevisionStoreReader {
    /// <summary>Reads the root revision-store structures from a file path.</summary>
    public static OneNoteRevisionStore Read(string path, OneNoteReaderOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));
        var file = new FileInfo(path);
        if (!file.Exists) throw new FileNotFoundException("The OneNote file does not exist.", path);
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
            return Read(stream, options);
        }
    }

    /// <summary>
    /// Reads the root revision-store structures from a caller-owned seekable stream and restores
    /// its original position.
    /// </summary>
    public static OneNoteRevisionStore Read(Stream stream, OneNoteReaderOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("Revision-store reading currently requires a seekable stream.", nameof(stream));
        var effectiveOptions = options ?? new OneNoteReaderOptions();
        effectiveOptions.Validate();
        if (effectiveOptions.MaxInputBytes.HasValue && stream.Length > effectiveOptions.MaxInputBytes.Value) {
            throw new IOException("OneNote input exceeds MaxInputBytes.");
        }

        long originalPosition = stream.Position;
        try {
            OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(stream, effectiveOptions);
            if (header.StorageFormat == OneNoteStorageFormat.FileSynchronizationPackage) {
                return OneNotePackageStoreReader.Read(stream, header, effectiveOptions);
            }
            if (header.StorageFormat != OneNoteStorageFormat.RevisionStore) {
                throw new OneNoteFormatException("ONENOTE_NOT_REVISION_STORE", "The OneNote artifact does not use a supported MS-ONESTORE encoding.");
            }
            if (!header.ExpectedFileLength.HasValue || !header.RootFileNodeList.HasValue) {
                throw new OneNoteFormatException("ONENOTE_REVISION_STORE_HEADER", "The revision-store header does not expose its required root structures.");
            }
            IReadOnlyDictionary<uint, int> committedNodeCounts = OneNoteTransactionLogReader.Read(
                stream,
                header,
                effectiveOptions);
            OneNoteFileNodeList root = OneNoteFileNodeListReader.Read(
                stream,
                header.RootFileNodeList.Value,
                header.ExpectedFileLength.Value,
                committedNodeCounts,
                effectiveOptions);
            ValidateRootList(root, header.FileKind);
            IReadOnlyList<OneNoteFileNodeList> lists = ReadReachableLists(
                stream,
                root,
                header.RootFileNodeList.Value.Offset,
                header.ExpectedFileLength.Value,
                committedNodeCounts,
                effectiveOptions);
            OneNoteRevisionStoreObjectReadResult objects = OneNoteRevisionStoreObjectReader.Read(
                stream,
                root,
                header.ExpectedFileLength.Value,
                effectiveOptions);
            return new OneNoteRevisionStore(
                header,
                root,
                lists,
                objects.Revisions.AsReadOnly(),
                objects.Objects.AsReadOnly(),
                objects.FileDataObjects.AsReadOnly());
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static IReadOnlyList<OneNoteFileNodeList> ReadReachableLists(
        Stream stream,
        OneNoteFileNodeList root,
        ulong rootOffset,
        ulong declaredFileLength,
        IReadOnlyDictionary<uint, int> committedNodeCounts,
        OneNoteReaderOptions options) {
        var lists = new List<OneNoteFileNodeList> { root };
        var byOffset = new Dictionary<ulong, OneNoteFileNodeList> { [rootOffset] = root };
        var queue = new Queue<OneNoteFileNodeList>();
        queue.Enqueue(root);
        int totalNodes = root.Nodes.Count;
        int totalFragments = root.Fragments.Count;

        while (queue.Count > 0) {
            OneNoteFileNodeList parent = queue.Dequeue();
            foreach (OneNoteFileNode node in parent.Nodes) {
                if (node.BaseType != OneNoteFileNodeBaseType.FileNodeListReference ||
                    node.ChunkReference == null ||
                    node.ChunkReference.IsNil ||
                    node.ChunkReference.IsZero) {
                    continue;
                }

                ulong childOffset = node.ChunkReference.Offset;
                if (byOffset.TryGetValue(childOffset, out OneNoteFileNodeList? existing)) {
                    node.ReferencedFileNodeList = existing;
                    continue;
                }
                if (node.ChunkReference.Length > uint.MaxValue) {
                    throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIST_SIZE", "A referenced file-node list fragment exceeds the desktop format's 32-bit length range.", node.FileOffset);
                }
                var firstFragment = new OneNoteFileChunkReference(childOffset, (uint)node.ChunkReference.Length);
                OneNoteFileNodeList child = OneNoteFileNodeListReader.Read(stream, firstFragment, declaredFileLength, committedNodeCounts, options);
                if (totalNodes > options.MaxFileNodes - child.Nodes.Count) {
                    throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIMIT", "The file-node limit was exceeded while traversing referenced lists.", node.FileOffset);
                }
                if (totalFragments > options.MaxFileNodeListFragments - child.Fragments.Count) {
                    throw new OneNoteFormatException("ONENOTE_FILE_NODE_FRAGMENT_LIMIT", "The file-node-list fragment limit was exceeded while traversing referenced lists.", node.FileOffset);
                }
                totalNodes += child.Nodes.Count;
                totalFragments += child.Fragments.Count;
                node.ReferencedFileNodeList = child;
                byOffset.Add(childOffset, child);
                lists.Add(child);
                queue.Enqueue(child);
            }
        }
        return lists.AsReadOnly();
    }

    internal static void ValidateRootList(OneNoteFileNodeList root, OneNoteFileKind fileKind) {
        int manifestReferences = root.Nodes.Count(node => node.RawId == (ushort)OneNoteFileNodeId.ObjectSpaceManifestListReference);
        int rootDeclarations = root.Nodes.Count(node => node.RawId == (ushort)OneNoteFileNodeId.ObjectSpaceManifestRoot);
        if (manifestReferences < 1 || rootDeclarations != 1) {
            throw new OneNoteFormatException("ONENOTE_ROOT_FILE_NODE_LIST", "The root file-node list does not contain the required object-space references and single root declaration.");
        }
        foreach (OneNoteFileNode node in root.Nodes) {
            bool allowed = node.RawId == (ushort)OneNoteFileNodeId.ObjectSpaceManifestListReference ||
                           node.RawId == (ushort)OneNoteFileNodeId.ObjectSpaceManifestRoot ||
                           node.RawId == (ushort)OneNoteFileNodeId.ChunkTerminator ||
                           (fileKind == OneNoteFileKind.Section && node.RawId == (ushort)OneNoteFileNodeId.FileDataStoreListReference);
            if (!allowed) {
                throw new OneNoteFormatException("ONENOTE_ROOT_FILE_NODE_TYPE", "The root file-node list contains a file-node type that is not valid at the root.", node.FileOffset);
            }
        }
    }
}
