using System.Text.Json.Serialization;

namespace OfficeIMO;

/// <summary>Portable conversion selection, output policy and bounded execution settings for one batch.</summary>
public sealed partial record OfficeBatchOptions {
    /// <summary>Optional display name retained in saved jobs.</summary>
    public string? Name { get; set; }
    /// <summary>Source directory discovered incrementally without following links.</summary>
    public string? InputDirectory { get; set; }
    /// <summary>Explicit local files, retaining paths relative to InputDirectory when supplied.</summary>
    public string[]? InputPaths { get; set; }
    /// <summary>Destination tree. Source and output trees must not overlap.</summary>
    public string OutputDirectory { get; set; } = string.Empty;
    /// <summary>Physical destination extension replacing the final source extension. Relative folders are retained when InputDirectory is supplied. Defaults to PDF.</summary>
    public string TargetExtension { get; set; } = ".pdf";
    /// <summary>Optional exact canonical route identifier.</summary>
    public string? ConversionRouteId { get; set; }
    /// <summary>Format-owned settings selected independently for each route.</summary>
    [JsonIgnore]
    public OfficeConversionOptions ConversionOptions { get; set; } = new();
    /// <summary>Cross-format output intent. Explicit adapter settings take precedence.</summary>
    public OfficeConversionProfile OutputProfile { get; set; } = OfficeConversionProfile.Faithful;
    /// <summary>File publication behavior. Durable checkpoint execution requires Fail.</summary>
    public OfficePublicationPolicy ConflictPolicy { get; set; } = OfficePublicationPolicy.Fail;
    /// <summary>Skip eligible inputs whose regular-file destination already exists, without claiming checkpoint reuse.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool SkipExistingOutputs { get; set; }
    /// <summary>Skip regular-file inputs whose regular-file output has a strictly newer UTC modification time.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool SkipUpToDate { get; set; }
    /// <summary>Whether directory discovery includes descendants.</summary>
    public bool Recursive { get; set; } = true;
    /// <summary>Optional source-extension filter, including leading dots.</summary>
    public string[]? Extensions { get; set; }
    /// <summary>Checkpoint manifest path. Private receipts and its lease live beside it in a .state directory.</summary>
    public string? ResumeFrom { get; set; }
    /// <summary>Maximum simultaneous files, from one through 32.</summary>
    public int MaxParallelism { get; set; } = 2;
    /// <summary>Maximum discovered files, including skipped inputs.</summary>
    public int MaximumFiles { get; set; } = 1_000_000;
    /// <summary>Maximum source bytes per file.</summary>
    public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum serialized output bytes per file.</summary>
    public long MaximumOutputBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum XML characters per package part.</summary>
    public long MaximumXmlCharactersInPart { get; set; } = 10L * 1024 * 1024;
    /// <summary>Retries recorded failures. Completed entries still require hash verification.</summary>
    public bool RetryFailed { get; set; }
    /// <summary>Admitted environment-variable reference for a source password. Its value is never saved.</summary>
    public string? SourcePasswordEnvironmentVariable { get; set; }
    /// <summary>Separate admitted environment-variable reference for PDF source passwords in mixed-format jobs.</summary>
    public string? PdfPasswordEnvironmentVariable { get; set; }
    /// <summary>Runtime resolver for explicitly admitted secret references.</summary>
    [JsonIgnore]
    public Func<string, string?>? ResolveEnvironmentVariable { get; set; }
    /// <summary>Optional new run report path. Checkpoint receipts remain the sole reuse authority.</summary>
    public string? ReportPath { get; set; }
    /// <summary>Whether an intentionally empty selection counts as a successful attempt.</summary>
    public bool AllowEmpty { get; set; }
    /// <summary>Host publication ownership policy, checked before staging and final publication.</summary>
    [JsonIgnore]
    public IOfficePublicationGuard? PublicationGuard { get; set; }
    /// <summary>Optional host retention policy for recorded reports.</summary>
    [JsonIgnore]
    public Func<bool>? RetainReport { get; set; }

    internal string? CheckpointDirectory => ResumeFrom == null ? null : ResumeFrom + ".state";
    internal string? JobPath { get; set; }
    internal string? JobSha256 { get; set; }
}
