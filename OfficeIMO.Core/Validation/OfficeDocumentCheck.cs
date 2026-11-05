namespace OfficeIMO;

/// <summary>Read-only checks performed by the format's native validation engine.</summary>
public enum OfficeDocumentCheck {
    /// <summary>Checks document structure and native operation readiness.</summary>
    Structure,
    /// <summary>Checks the format schema when a native schema validator is available.</summary>
    Schema,
    /// <summary>Inspects package security without executing active content.</summary>
    Security,
    /// <summary>Runs the format's accessibility inspector.</summary>
    Accessibility,
    /// <summary>Assesses the requested formal compliance profile.</summary>
    Compliance
}
