namespace OfficeIMO.OneNote;

/// <summary>
/// A decoded MS-FSSHTTPB extended GUID.
/// </summary>
public sealed class OneNoteExtendedGuid {
    internal OneNoteExtendedGuid(Guid identifier, uint value, int encodedLength) {
        Identifier = identifier;
        Value = value;
        EncodedLength = encodedLength;
    }

    /// <summary>The GUID component.</summary>
    public Guid Identifier { get; }

    /// <summary>The unsigned integer component.</summary>
    public uint Value { get; }

    /// <summary>Number of encoded bytes consumed.</summary>
    public int EncodedLength { get; }

    /// <inheritdoc />
    public override string ToString() => "{{" + Identifier.ToString("D") + "}," + Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + "}";

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OneNoteExtendedGuid other && Identifier == other.Identifier && Value == other.Value;

    /// <inheritdoc />
    public override int GetHashCode() => (Identifier.GetHashCode() * 397) ^ Value.GetHashCode();
}

/// <summary>
/// A file offset and byte count from an MS-ONESTORE header.
/// </summary>
public readonly struct OneNoteFileChunkReference {
    internal OneNoteFileChunkReference(ulong offset, uint length) {
        Offset = offset;
        Length = length;
    }

    /// <summary>Unsigned byte offset from the start of the file.</summary>
    public ulong Offset { get; }

    /// <summary>Referenced byte count.</summary>
    public uint Length { get; }

    /// <summary>True when both the offset and byte count are zero.</summary>
    public bool IsZero => Offset == 0 && Length == 0;

    /// <summary>True when the offset is all bits set and the byte count is zero.</summary>
    public bool IsNil => Offset == ulong.MaxValue && Length == 0;

    /// <inheritdoc />
    public override string ToString() => IsNil ? "nil" : Offset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "+" + Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
}

/// <summary>
/// Decoded physical header for a <c>.one</c> or <c>.onetoc2</c> file.
/// </summary>
public sealed class OneNoteFileHeader {
    private readonly List<OneNoteDiagnostic> _diagnostics = new List<OneNoteDiagnostic>();

    /// <summary>Logical artifact kind.</summary>
    public OneNoteFileKind FileKind { get; internal set; }

    /// <summary>Physical storage encoding.</summary>
    public OneNoteStorageFormat StorageFormat { get; internal set; }

    /// <summary>MS-ONESTORE file-type identifier.</summary>
    public Guid FileTypeId { get; internal set; }

    /// <summary>Identity of this file.</summary>
    public Guid FileId { get; internal set; }

    /// <summary>Legacy version or package version identifier.</summary>
    public Guid LegacyFileVersionId { get; internal set; }

    /// <summary>Identifier of the physical file encoding.</summary>
    public Guid FileFormatId { get; internal set; }

    /// <summary>Actual source length when available.</summary>
    public long? ActualFileLength { get; internal set; }

    /// <summary>Expected file length declared by a desktop revision-store header.</summary>
    public ulong? ExpectedFileLength { get; internal set; }

    /// <summary>Number of complete transactions declared by a desktop revision-store header.</summary>
    public uint? TransactionCount { get; internal set; }

    /// <summary>Table-of-contents ancestor identity declared by a desktop revision-store header.</summary>
    public Guid? AncestorId { get; internal set; }

    /// <summary>Identity of the current desktop file version.</summary>
    public Guid? FileVersionId { get; internal set; }

    /// <summary>Generation number of the current desktop file version.</summary>
    public ulong? FileVersionGeneration { get; internal set; }

    /// <summary>Identity that changes while existing desktop file contents are rewritten.</summary>
    public Guid? DenyReadFileVersionId { get; internal set; }

    /// <summary>Hashed chunk list reference, when present.</summary>
    public OneNoteFileChunkReference? HashedChunkList { get; internal set; }

    /// <summary>Transaction log reference for a desktop revision store.</summary>
    public OneNoteFileChunkReference? TransactionLog { get; internal set; }

    /// <summary>Root file-node-list reference for a desktop revision store.</summary>
    public OneNoteFileChunkReference? RootFileNodeList { get; internal set; }

    /// <summary>Free chunk list reference, when present.</summary>
    public OneNoteFileChunkReference? FreeChunkList { get; internal set; }

    /// <summary>Storage index identity for an MS-FSSHTTPB package-store file.</summary>
    public OneNoteExtendedGuid? StorageIndexId { get; internal set; }

    /// <summary>Cell schema that distinguishes packaged <c>.one</c> and <c>.onetoc2</c> data.</summary>
    public Guid? CellSchemaId { get; internal set; }

    /// <summary>Non-fatal compatibility and fidelity diagnostics.</summary>
    public IReadOnlyList<OneNoteDiagnostic> Diagnostics => _diagnostics;

    internal void AddDiagnostic(string code, string message, long? offset = null) {
        _diagnostics.Add(new OneNoteDiagnostic {
            Code = code,
            Severity = OneNoteDiagnosticSeverity.Warning,
            Message = message,
            Offset = offset
        });
    }
}
