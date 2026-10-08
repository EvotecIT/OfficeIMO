namespace OfficeIMO.Access;

/// <summary>Structured field kinds. Unknown structures remain accessible as typed native backing rows.</summary>
public enum AccessComplexKind {
    /// <summary>Several values of one scalar type.</summary>
    MultiValue,
    /// <summary>Embedded files with metadata and encoded native payloads.</summary>
    Attachment,
    /// <summary>Append-only text versions and their timestamps.</summary>
    VersionHistory,
    /// <summary>A structure without a qualified convenience decoder.</summary>
    Unknown
}

/// <summary>Native structured-field schema. Backing storage stays separate from ordinary user tables.</summary>
public sealed class AccessComplexDefinition {
    internal AccessComplexDefinition(AccessNativeTable table, int foreignOrdinal, AccessComplexKind kind, AccessColumn[] fields) { Table = table; ForeignOrdinal = foreignOrdinal; Kind = kind; Fields = Array.AsReadOnly(fields); }
    internal AccessNativeTable Table { get; }
    internal int ForeignOrdinal { get; }
    /// <summary>Qualified structured-field kind.</summary>
    public AccessComplexKind Kind { get; }
    /// <summary>Value/attachment fields, excluding internal row and parent keys.</summary>
    public IReadOnlyList<AccessColumn> Fields { get; }
}

/// <summary>A lazy view of the structured values belonging to one parent record. The document must remain open.</summary>
public sealed class AccessComplexValue {
    internal AccessComplexValue(AccessComplexDefinition definition, int key) { Definition = definition; ForeignKey = key; }
    /// <summary>Structured-field schema.</summary>
    public AccessComplexDefinition Definition { get; }
    /// <summary>Stored parent value key; no provider or expression resolves it.</summary>
    public int ForeignKey { get; }
    /// <summary>Opens forward-only backing rows filtered by this parent key. Large fields decode only when requested.</summary>
    public AccessDataReader OpenDataReader(CancellationToken cancellationToken = default) {
        Definition.Table.Database.Document.EnsureNotDisposed(); cancellationToken.ThrowIfCancellationRequested();
        var cursor = new AccessNativeRowCursor(Definition.Table, cancellationToken, Definition.ForeignOrdinal, ForeignKey);
        try { return new AccessDataReader(Definition.Table.Model!, cancellationToken, cursor); }
        catch { cursor.Dispose(); throw; }
    }
    /// <summary>Enumerates scalar values for a multivalued field. Null and empty values retain their original meaning.</summary>
    public IEnumerable<object?> EnumerateValues(CancellationToken cancellationToken = default) {
        if (Definition.Kind != AccessComplexKind.MultiValue || Definition.Fields.Count != 1) throw new InvalidOperationException("This structured field does not have one multivalued scalar definition.");
        using var rows = OpenDataReader(cancellationToken); int ordinal = rows.GetOrdinal(Definition.Fields[0].Name);
        while (rows.Read()) yield return rows.IsDBNull(ordinal) ? null : rows.GetValue(ordinal);
    }
    /// <summary>Enumerates attachment metadata. File content remains lazy until GetBytes/GetEncodedBytes is requested.</summary>
    public IEnumerable<AccessAttachment> EnumerateAttachments(CancellationToken cancellationToken = default) {
        if (Definition.Kind != AccessComplexKind.Attachment) throw new InvalidOperationException("This structured field is not an attachment field.");
        var table = Definition.Table; table.Database.Document.EnsureNotDisposed();
        using var rows = new AccessNativeRowCursor(table, cancellationToken, Definition.ForeignOrdinal, ForeignKey);
        while (rows.Read(cancellationToken)) yield return new AccessAttachment(table, rows.Current, cancellationToken);
    }
}
