namespace OfficeIMO.Access;

/// <summary>A stored reference observed without resolving files, queries, expressions or code.</summary>
public sealed class AccessDependency {
    internal AccessDependency(Guid source, string kind, string reference, Guid? target) { SourceObjectId = source; Kind = kind; Reference = reference; TargetObjectId = target; }
    /// <summary>Model identity of the source object.</summary>
    public Guid SourceObjectId { get; }
    /// <summary>Reference role, such as query-table or record-source.</summary>
    public string Kind { get; }
    /// <summary>Stored name or expression; it remains inert.</summary>
    public string Reference { get; }
    /// <summary>Unambiguous local object identity for an exact name reference. SQL expressions are never guessed as resolved names.</summary>
    public Guid? TargetObjectId { get; }
}

/// <summary>A model mutation recorded independently from native page allocation or application execution.</summary>
public sealed class AccessChange {
    internal AccessChange(long revision, Guid objectId, string operation) { Revision = revision; ObjectId = objectId; Operation = operation; }
    /// <summary>Document mutation revision.</summary>
    public long Revision { get; }
    /// <summary>Affected model identity.</summary>
    public Guid ObjectId { get; }
    /// <summary>Mutation kind. It does not imply an in-place native edit.</summary>
    public string Operation { get; }
}
