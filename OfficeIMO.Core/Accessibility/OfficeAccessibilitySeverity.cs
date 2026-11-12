namespace OfficeIMO;

/// <summary>Severity assigned to an accessibility finding.</summary>
public enum OfficeAccessibilitySeverity {
    /// <summary>Useful inspection context.</summary>
    Info,
    /// <summary>Potential accessibility risk that deserves review.</summary>
    Warning,
    /// <summary>Policy violation that should fail the selected CI gate.</summary>
    Error
}

/// <summary>Built-in accessibility policy profiles shared by the Word and PowerPoint checkers.</summary>
public enum OfficeAccessibilityPolicyProfile {
    /// <summary>Practical policy suitable for ordinary CI gates. Structural risks are warnings.</summary>
    Default,
    /// <summary>Stricter policy: structural risks are errors and explicit metadata is required. Heuristic findings stay warnings.</summary>
    Strict
}
