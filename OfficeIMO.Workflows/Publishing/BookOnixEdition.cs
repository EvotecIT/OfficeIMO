namespace OfficeIMO.Workflows;

/// <summary>Explicit publication edition characteristics from ONIX list 21.</summary>
public enum BookOnixEditionType {
    /// <summary>Abridged edition (ABR).</summary>
    Abridged,
    /// <summary>Unabridged edition (UBR).</summary>
    Unabridged,
    /// <summary>Annotated edition (ANN).</summary>
    Annotated,
    /// <summary>Revised edition (REV).</summary>
    Revised,
    /// <summary>Enlarged or expanded edition (ENL).</summary>
    Enlarged,
    /// <summary>Illustrated edition (ILL).</summary>
    Illustrated,
    /// <summary>Critical edition (CRI).</summary>
    Critical,
    /// <summary>New edition (NED).</summary>
    New
}

/// <summary>A complete plain-text edition description for display, with an optional ONIX list 74 language.</summary>
public sealed record BookOnixEditionStatement(string Text, string? LanguageCode = null);

/// <summary>Publisher-supplied edition information, independent of internal project revision history.</summary>
public sealed record BookOnixEdition {
    /// <summary>Explicitly asserts that no edition information applies. Mutually exclusive with all other fields.</summary>
    public bool NoEdition { get; init; }
    /// <summary>At most eight distinct edition characteristics.</summary>
    public IReadOnlyList<BookOnixEditionType> Types { get; init; } = [];
    /// <summary>Positive edition number, when numbered.</summary>
    public int? Number { get; init; }
    /// <summary>Minor revision within the numbered edition. Requires Number; never inferred from project revisions.</summary>
    public string? VersionNumber { get; init; }
    /// <summary>At most 16 complete plain-text descriptions, each with a distinct language (including unspecified).</summary>
    public IReadOnlyList<BookOnixEditionStatement> Statements { get; init; } = [];
}
