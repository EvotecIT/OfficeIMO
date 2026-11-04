namespace OfficeIMO;

/// <summary>Lifecycle of a recorded batch attempt, independent of checkpoint integrity.</summary>
public enum OfficeBatchRunState {
    /// <summary>The attempt holds its report lease.</summary>
    Running,
    /// <summary>The attempt completed successfully.</summary>
    Completed,
    /// <summary>The attempt encountered conversion or discovery failures.</summary>
    Failed,
    /// <summary>The caller cancelled the attempt.</summary>
    Cancelled,
    /// <summary>No final record or live writer lease remains.</summary>
    Interrupted
}
