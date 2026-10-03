namespace OfficeIMO.IWork;

/// <summary>Why a declared reference was not resolved for an assessed content path.</summary>
public enum IWorkSourceReferenceIssueKind {
    /// <summary>The reference identifies an object that is absent from the source index.</summary>
    MissingTarget,
    /// <summary>The declared value is not a valid, unambiguous native object reference.</summary>
    MalformedReference,
    /// <summary>The reader rejected the containing reference set, so this otherwise readable reference was not selected.</summary>
    RejectedReferenceSet,
    /// <summary>The referenced object exists, but its native message type is not supported by the selected content path.</summary>
    UnexpectedTargetType
}

/// <summary>One unresolved declared source-field occurrence, reported once even when its owner is selected repeatedly.</summary>
public sealed class IWorkSourceReferenceIssue {
    internal IWorkSourceReferenceIssue(IWorkObjectIdentity owner, string fieldPath, int referenceIndex,
        ulong? targetIdentifier, IWorkSourceReferenceIssueKind kind) {
        Owner = owner;
        FieldPath = fieldPath;
        ReferenceIndex = referenceIndex;
        TargetIdentifier = targetIdentifier;
        Kind = kind;
    }

    /// <summary>Gets the source record containing the assessed reference field.</summary>
    public IWorkObjectIdentity Owner { get; }
    /// <summary>Gets the protobuf field path from the owner payload. Nested repeated messages use one-based bracket indexes.</summary>
    public string FieldPath { get; }
    /// <summary>Gets the one-based occurrence within the assessed reference field. Repeated references are counted separately.</summary>
    public int ReferenceIndex { get; }
    /// <summary>Gets the native target identifier when it can be read unambiguously; this does not establish that a target exists.</summary>
    public ulong? TargetIdentifier { get; }
    /// <summary>Gets why the declared reference occurrence was not resolved.</summary>
    public IWorkSourceReferenceIssueKind Kind { get; }
}
