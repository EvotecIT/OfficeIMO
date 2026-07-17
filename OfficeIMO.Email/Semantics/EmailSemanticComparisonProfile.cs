namespace OfficeIMO.Email;

/// <summary>Defines which representation details participate in semantic email comparison.</summary>
public enum EmailSemanticComparisonProfile {
    /// <summary>
    /// Compares portable message, Outlook, recipient, attachment, body, and extended-MAPI semantics while
    /// normalizing store-generated identifiers and serialization details.
    /// </summary>
    Migration = 0,

    /// <summary>
    /// Adds source-format, raw-header, raw-MAPI, TNEF, property-bag, and MIME representation details.
    /// </summary>
    Strict = 1,

    /// <summary>
    /// Compares portable content while excluding store identity, access, synchronization, and modification stamps
    /// that commonly differ when the same logical item appears in multiple stores.
    /// </summary>
    Deduplication = 2
}
