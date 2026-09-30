namespace OfficeIMO.IWork;

/// <summary>Identifies a source object without retaining its protobuf payload.</summary>
public sealed class IWorkObjectIdentity {
    internal IWorkObjectIdentity(IWorkArchiveRecord record) {
        RecordIdentifier = record.Identifier;
        MessageType = record.MessageType;
        EntryPath = record.EntryPath;
        PayloadIndex = record.PayloadIndex;
    }

    /// <summary>Gets the native IWA object identifier.</summary>
    public ulong RecordIdentifier { get; }
    /// <summary>Gets the registry message type.</summary>
    public uint MessageType { get; }
    /// <summary>Gets the normalized IWA package entry path.</summary>
    public string EntryPath { get; }
    /// <summary>Gets the payload position within its ArchiveInfo group.</summary>
    public int PayloadIndex { get; }
}

/// <summary>The selected primary-record units counted independently of paragraphs or cells.</summary>
public enum IWorkSourceUnitKind {
    /// <summary>A document root.</summary>
    Document,
    /// <summary>A Numbers sheet.</summary>
    Sheet,
    /// <summary>A Keynote slide.</summary>
    Slide,
    /// <summary>A table-info drawable, excluding its auxiliary model and tiles.</summary>
    Table,
    /// <summary>A text storage, counted once even when used by multiple drawables.</summary>
    Text,
    /// <summary>An image drawable, counted independently of a shared resource.</summary>
    Image,
    /// <summary>A selected object whose native type has no supported semantic reconstruction.</summary>
    UnsupportedObject
}

/// <summary>The known destination outcome of one recognized source unit.</summary>
public enum IWorkSourceUnitDisposition {
    /// <summary>The unit has an editable destination representation; individual fields may still have fidelity loss.</summary>
    Reconstructed,
    /// <summary>The unit was explicitly excluded from the editable destination.</summary>
    Omitted,
    /// <summary>The selected unit's destination representation has not been established, including visual fallback.</summary>
    Unassessed
}

/// <summary>Identity and destination accounting for one selected primary source record.</summary>
public sealed class IWorkSourceUnit {
    internal IWorkSourceUnit(IWorkSourceUnitKind kind, IWorkObjectIdentity identity, IWorkSourceUnitDisposition disposition) {
        Kind = kind;
        Identity = identity;
        Disposition = disposition;
    }
    /// <summary>Gets the unit kind.</summary>
    public IWorkSourceUnitKind Kind { get; }
    /// <summary>Gets its native object identity.</summary>
    public IWorkObjectIdentity Identity { get; }
    /// <summary>Gets its known editable destination outcome.</summary>
    public IWorkSourceUnitDisposition Disposition { get; }
}

/// <summary>Counts identified units selected by the semantic projection, including explicitly omitted unsupported objects. Inactive records and unresolved references are not included.</summary>
public sealed class IWorkSourceUnitCount {
    internal IWorkSourceUnitCount(IWorkSourceUnitKind kind, IEnumerable<IWorkSourceUnit> units) {
        Kind = kind;
        foreach (IWorkSourceUnit unit in units) {
            if (unit.Kind != kind) continue;
            TotalCount++;
            if (unit.Disposition == IWorkSourceUnitDisposition.Reconstructed) ReconstructedCount++;
            else if (unit.Disposition == IWorkSourceUnitDisposition.Omitted) OmittedCount++;
            else UnassessedCount++;
        }
    }
    /// <summary>Gets the unit kind.</summary>
    public IWorkSourceUnitKind Kind { get; }
    /// <summary>Gets the identified selected-unit count.</summary>
    public int TotalCount { get; }
    /// <summary>Gets the count with an editable destination representation.</summary>
    public int ReconstructedCount { get; }
    /// <summary>Gets the count explicitly excluded from the editable destination.</summary>
    public int OmittedCount { get; }
    /// <summary>Gets the count whose destination representation is unassessed.</summary>
    public int UnassessedCount { get; }
}
