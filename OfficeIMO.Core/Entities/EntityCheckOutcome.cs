namespace OfficeIMO.Entities;

/// <summary>Outcome of a scheme’s checksum.</summary>
public enum EntityCheckOutcome {
    /// <summary>No check was applied.</summary>
    NotDefined,
    /// <summary>The check passed.</summary>
    Passed,
    /// <summary>The check failed.</summary>
    Failed,
}
