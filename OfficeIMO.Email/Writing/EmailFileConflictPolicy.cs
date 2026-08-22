namespace OfficeIMO.Email;

/// <summary>
/// Controls how a path-based email save handles an existing destination.
/// </summary>
public enum EmailFileConflictPolicy {
    /// <summary>Fails without modifying the existing destination.</summary>
    FailIfExists,
    /// <summary>Replaces the existing destination after the new artifact has been written successfully.</summary>
    Replace
}
