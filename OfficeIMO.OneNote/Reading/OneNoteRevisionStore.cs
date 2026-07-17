namespace OfficeIMO.OneNote;

/// <summary>A physical file-node-list fragment from an MS-ONESTORE revision store.</summary>
public sealed class OneNoteFileNodeListFragment {
    internal OneNoteFileNodeListFragment(
        long fileOffset,
        int length,
        uint fileNodeListId,
        uint sequence,
        OneNoteFileChunkReference nextFragment,
        IReadOnlyList<OneNoteFileNode> nodes) {
        FileOffset = fileOffset;
        Length = length;
        FileNodeListId = fileNodeListId;
        Sequence = sequence;
        NextFragment = nextFragment;
        Nodes = nodes;
    }

    /// <summary>Absolute byte offset of this fragment.</summary>
    public long FileOffset { get; }

    /// <summary>Encoded fragment length.</summary>
    public int Length { get; }

    /// <summary>Identity shared by every fragment in this list.</summary>
    public uint FileNodeListId { get; }

    /// <summary>Zero-based sequence number within the list.</summary>
    public uint Sequence { get; }

    /// <summary>Reference to the next fragment, or the nil sentinel for the last fragment.</summary>
    public OneNoteFileChunkReference NextFragment { get; }

    /// <summary>File nodes decoded from this fragment.</summary>
    public IReadOnlyList<OneNoteFileNode> Nodes { get; }
}

/// <summary>A complete chain of MS-ONESTORE file-node-list fragments.</summary>
public sealed class OneNoteFileNodeList {
    internal OneNoteFileNodeList(uint id, IReadOnlyList<OneNoteFileNodeListFragment> fragments, IReadOnlyList<OneNoteFileNode> nodes) {
        Id = id;
        Fragments = fragments;
        Nodes = nodes;
    }

    /// <summary>File-node-list identity.</summary>
    public uint Id { get; }

    /// <summary>Fragments in sequence order.</summary>
    public IReadOnlyList<OneNoteFileNodeListFragment> Fragments { get; }

    /// <summary>All nodes in logical list order.</summary>
    public IReadOnlyList<OneNoteFileNode> Nodes { get; }
}

/// <summary>Decoded physical root of a desktop MS-ONESTORE revision store.</summary>
public sealed class OneNoteRevisionStore {
    internal OneNoteRevisionStore(
        OneNoteFileHeader header,
        OneNoteFileNodeList rootFileNodeList,
        IReadOnlyList<OneNoteFileNodeList> fileNodeLists,
        IReadOnlyList<OneNoteRevisionManifest> revisions,
        IReadOnlyList<OneNoteRevisionStoreObject> objects,
        IReadOnlyList<OneNoteFileDataStoreObject> fileDataObjects) {
        Header = header;
        RootFileNodeList = rootFileNodeList;
        FileNodeLists = fileNodeLists;
        Revisions = revisions;
        Objects = objects;
        FileDataObjects = fileDataObjects;
    }

    /// <summary>Validated revision-store header.</summary>
    public OneNoteFileHeader Header { get; }

    /// <summary>Root file-node list that identifies all object spaces in the file.</summary>
    public OneNoteFileNodeList RootFileNodeList { get; }

    /// <summary>All reachable file-node lists, beginning with <see cref="RootFileNodeList"/>.</summary>
    public IReadOnlyList<OneNoteFileNodeList> FileNodeLists { get; }

    /// <summary>Reachable revision manifests in source order.</summary>
    public IReadOnlyList<OneNoteRevisionManifest> Revisions { get; }

    /// <summary>Decoded object declarations and revisions in source order.</summary>
    public IReadOnlyList<OneNoteRevisionStoreObject> Objects { get; }

    /// <summary>Embedded file-data store payloads.</summary>
    public IReadOnlyList<OneNoteFileDataStoreObject> FileDataObjects { get; }
}
