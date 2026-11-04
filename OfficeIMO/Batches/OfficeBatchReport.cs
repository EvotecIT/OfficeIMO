namespace OfficeIMO;

/// <summary>Terminal state of one selected batch item.</summary>
public enum OfficeBatchItemStatus {
    /// <summary>Conversion and publication completed.</summary>
    Completed,
    /// <summary>Caller cancellation prevented completion.</summary>
    Cancelled,
    /// <summary>Conversion or publication failed.</summary>
    Failed,
    /// <summary>A provider write began but its final contents could not be verified.</summary>
    Unconfirmed
}

/// <summary>Complete bounded counts, independent of retained item observations.</summary>
public sealed record OfficeBatchSummary(long Selected, long Completed, long Reused, long Failed, bool Cancelled, long Skipped = 0) {
    /// <summary>Terminal discovery or execution-budget failure; completed artifacts remain available.</summary>
    public string? FailureSummary { get; init; }
}

/// <summary>One conversion outcome. Skipped inputs have no conversion status or output.</summary>
public sealed record OfficeBatchItemReport(string InputPath, string? OutputPath, OfficeBatchItemStatus? Status,
    bool Reused, string Summary, IReadOnlyList<OfficeDiagnostic> Diagnostics, bool Skipped = false) {
    /// <summary>OutputExists or UpToDate when an eligible input was intentionally skipped; null for unselected inputs.</summary>
    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
    public string? SkippedReason { get; init; }
    /// <summary>Completed stages, when the conversion produced a report.</summary>
    [System.Text.Json.Serialization.JsonConverter(typeof(OfficeBatchConversionReportConverter))]
    public OfficeConversionReport? Report { get; init; }
}

/// <summary>Bounded observations and complete counts from one batch attempt.</summary>
public sealed partial record OfficeBatchReport {
    /// <summary>Versioned report contract.</summary>
    public const string CurrentSchema = "officeimo.conversion.batch.v2";
    /// <summary>Report contract identifier.</summary>
    public string Schema { get; init; } = CurrentSchema;
    /// <summary>Whether execution completed without failures or cancellation under its empty-selection policy.</summary>
    public bool Succeeded { get; init; }
    /// <summary>Complete aggregate counts.</summary>
    public required OfficeBatchSummary Summary { get; init; }
    /// <summary>Bounded item observations, in callback order.</summary>
    public IReadOnlyList<OfficeBatchItemReport> Items { get; init; } = [];
    /// <summary>Whether observations were omitted to enforce report limits.</summary>
    public bool ItemsTruncated { get; init; }
    /// <summary>Unique recorded-attempt identifier, when run recording was requested.</summary>
    public string? RunId { get; init; }
    /// <summary>Attempt lifecycle; unfinished saved records recover as Interrupted when no writer lease remains.</summary>
    public OfficeBatchRunState State { get; init; } = OfficeBatchRunState.Completed;
    /// <summary>Absolute output directory for a recorded attempt.</summary>
    public string? OutputDirectory { get; init; }
    /// <summary>Checkpoint manifest used by the attempt.</summary>
    public string? ResumeFrom { get; init; }
    /// <summary>Whether this attempt retried prior failures.</summary>
    public bool RetryFailed { get; init; }
    /// <summary>Whether an empty selection was explicitly accepted.</summary>
    public bool AllowEmpty { get; init; }
    /// <summary>Unexpected execution failure, when ordinary aggregate results could not be returned.</summary>
    public string? FailureSummary { get; init; }
    /// <summary>Attempt start time.</summary>
    public DateTimeOffset StartedUtc { get; init; }
    /// <summary>Attempt finish time, or null while a recorded attempt is running.</summary>
    public DateTimeOffset? FinishedUtc { get; init; }
    /// <summary>Saved job path, when the attempt came from a saved definition.</summary>
    public string? JobPath { get; init; }
    /// <summary>Hash of the exact saved job bytes, including any byte-order mark.</summary>
    public string? JobSha256 { get; init; }
    /// <summary>Report publication failure. Conversion counts remain available.</summary>
    public string? RecordingError { get; init; }
}
