namespace OfficeIMO.IWork;

/// <summary>Why selected native declarations could not be assessed by the supported reader.</summary>
public enum IWorkSourceDeclarationIssueKind {
    /// <summary>A selected payload or nested message could not be decoded.</summary>
    MalformedMessage,
    /// <summary>The declared message set has an unsupported or ambiguous envelope.</summary>
    RejectedMessageSet,
    /// <summary>A decoded declaration has invalid metadata needed to select its content.</summary>
    InvalidSelectionMetadata,
    /// <summary>A decoded declaration contains an invalid or unsupported scalar value.</summary>
    InvalidValue
}

/// <summary>One selected source path whose declarations could not be assessed. It does not identify references or omitted objects.</summary>
public sealed class IWorkSourceDeclarationIssue {
    internal IWorkSourceDeclarationIssue(IWorkObjectIdentity owner, string fieldPath,
        int? declaredValueCount, IWorkSourceDeclarationIssueKind kind) {
        Owner = owner;
        FieldPath = fieldPath;
        DeclaredValueCount = declaredValueCount;
        Kind = kind;
    }

    /// <summary>Gets the record containing the selected declaration.</summary>
    public IWorkObjectIdentity Owner { get; }
    /// <summary>Gets the protobuf field path, with one-based repeated-message indexes. "$" identifies a whole record payload.</summary>
    public string FieldPath { get; }
    /// <summary>Gets the known outer field-value count, or null when no outer field count is available, including whole payloads. This is not the number of nested references, content objects or omitted items.</summary>
    public int? DeclaredValueCount { get; }
    /// <summary>Gets why the declarations at this path could not be assessed.</summary>
    public IWorkSourceDeclarationIssueKind Kind { get; }
}
