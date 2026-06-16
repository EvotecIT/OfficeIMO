namespace OfficeIMO.Rtf;

/// <summary>
/// Run-level revision kind represented by common RTF revision controls.
/// </summary>
public enum RtfRevisionKind {
    /// <summary>No tracked revision marker.</summary>
    None,

    /// <summary>Inserted text marked with <c>\revised</c>.</summary>
    Inserted,

    /// <summary>Deleted text marked with <c>\deleted</c>.</summary>
    Deleted
}
