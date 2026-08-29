namespace OfficeIMO.IWork;

/// <summary>One payload record preserved from an IWA object stream.</summary>
public sealed class IWorkArchiveRecord {
    private readonly byte[] _payload;

    internal IWorkArchiveRecord(ulong identifier, uint messageType, IReadOnlyList<uint> versions,
        IReadOnlyList<ulong> objectReferences, IReadOnlyList<ulong> dataReferences,
        string entryPath, int payloadIndex, byte[] payload) {
        Identifier = identifier;
        MessageType = messageType;
        Versions = versions;
        ObjectReferences = objectReferences;
        DataReferences = dataReferences;
        EntryPath = entryPath;
        PayloadIndex = payloadIndex;
        _payload = payload;
    }

    /// <summary>Gets the object identifier declared by ArchiveInfo.</summary>
    public ulong Identifier { get; }
    /// <summary>Gets the application-specific registry message type.</summary>
    public uint MessageType { get; }
    /// <summary>Gets the version vector declared by MessageInfo.</summary>
    public IReadOnlyList<uint> Versions { get; }
    /// <summary>Gets object identifiers referenced by MessageInfo.</summary>
    public IReadOnlyList<ulong> ObjectReferences { get; }
    /// <summary>Gets data identifiers referenced by MessageInfo.</summary>
    public IReadOnlyList<ulong> DataReferences { get; }
    /// <summary>Gets the IWA package entry containing this record.</summary>
    public string EntryPath { get; }
    /// <summary>Gets the zero-based payload position within the ArchiveInfo group.</summary>
    public int PayloadIndex { get; }
    /// <summary>Gets whether this is the primary object payload in its ArchiveInfo group.</summary>
    public bool IsPrimary => PayloadIndex == 0;
    /// <summary>Gets the raw payload length.</summary>
    public int PayloadLength => _payload.Length;
    /// <summary>Returns a defensive copy of the raw protobuf payload.</summary>
    public byte[] GetPayload() => (byte[])_payload.Clone();
    internal byte[] Payload => _payload;
}
