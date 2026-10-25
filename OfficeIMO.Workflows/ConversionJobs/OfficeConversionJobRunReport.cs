using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

/// <summary>The lifecycle of one recorded attempt, independent of checkpoint integrity.</summary>
public enum OfficeConversionJobRunState {
    /// <summary>The recorded attempt still holds its report lease.</summary>
    Running,
    /// <summary>Execution finished without item, discovery or cancellation failures.</summary>
    Completed,
    /// <summary>Execution encountered an item, discovery or unexpected failure.</summary>
    Failed,
    /// <summary>The caller cancelled execution.</summary>
    Cancelled,
    /// <summary>The attempt has no final record and no live writer lease.</summary>
    Interrupted
}

/// <summary>Observational history for one saved-job attempt. Its paths and outcomes never authorize output reuse.</summary>
public sealed record OfficeConversionJobRunReport {
    /// <summary>The versioned saved-job run contract.</summary>
    public const string CurrentSchema = "officeimo.conversion.job.run.v1";
    /// <summary>Versioned contract identifier.</summary>
    [JsonRequired]
    public string Schema { get; init; } = CurrentSchema;
    /// <summary>Unique attempt identifier, formatted as a GUID without separators.</summary>
    public required string RunId { get; init; }
    /// <summary>UTC instant when the attempt was admitted.</summary>
    public DateTimeOffset StartedUtc { get; init; }
    /// <summary>UTC instant when execution reached a terminal outcome, if recorded.</summary>
    public DateTimeOffset? FinishedUtc { get; init; }
    /// <summary>Absolute path to the admitted saved-job file.</summary>
    public required string JobPath { get; init; }
    /// <summary>SHA-256 of the exact admitted job-file bytes, including any byte-order mark.</summary>
    public required string JobSha256 { get; init; }
    /// <summary>Absolute conversion output directory.</summary>
    public required string OutputDirectory { get; init; }
    /// <summary>Optional absolute checkpoint directory. Checkpoints remain the sole reuse authority.</summary>
    public string? CheckpointDirectory { get; init; }
    /// <summary>Whether this attempt explicitly retries prior failed checkpoint items.</summary>
    public bool RetryFailed { get; init; }
    /// <summary>Whether an intentionally empty selection is accepted for this attempt.</summary>
    public bool AllowEmpty { get; init; }
    /// <summary>Attempt lifecycle; readers recover abandoned Running records as Interrupted.</summary>
    [JsonRequired]
    public OfficeConversionJobRunState State { get; init; }
    /// <summary>Bounded conversion observations and complete aggregate counts, available for ordinary returned outcomes.</summary>
    public OfficeConversionBatchReport? Batch { get; init; }
    /// <summary>Unexpected execution error, when execution did not produce an ordinary terminal result.</summary>
    public string? FailureSummary { get; init; }
    /// <summary>Report publication failure. Returned conversion counts remain available even if final persistence failed.</summary>
    public string? RecordingError { get; init; }
}
