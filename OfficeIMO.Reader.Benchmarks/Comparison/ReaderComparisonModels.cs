namespace OfficeIMO.Reader.Benchmarks.Comparison;

internal enum ReaderComparisonProbeKind {
    ContainsText,
    MarkdownHeading,
    MarkdownListItem,
    MarkdownTable,
    MarkdownLink,
    MarkdownImage,
    RichTable,
    RichLink,
    RichAsset,
    LocationPath,
    LocationHeading,
    LocationSheet,
    LocationSlide,
    LocationPage,
    RejectsMalformedInput
}

internal sealed class ReaderComparisonProbe {
    public ReaderComparisonProbe(
        string id,
        ReaderComparisonProbeKind kind,
        string marker = "",
        string expectedTarget = "") {
        Id = id;
        Kind = kind;
        Marker = marker;
        ExpectedTarget = expectedTarget;
    }

    public string Id { get; }
    public ReaderComparisonProbeKind Kind { get; }
    public string Marker { get; }
    public string ExpectedTarget { get; }
}

internal sealed class ReaderComparisonCase {
    public ReaderComparisonCase(
        string id,
        string sourceName,
        byte[] bytes,
        IReadOnlyList<ReaderComparisonProbe> probes) {
        Id = id;
        SourceName = sourceName;
        Bytes = bytes;
        Probes = probes;
    }

    public string Id { get; }
    public string SourceName { get; }
    public byte[] Bytes { get; }
    public IReadOnlyList<ReaderComparisonProbe> Probes { get; }
}

internal sealed class ReaderComparisonRunnerConfiguration {
    public string Name { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public List<string> Arguments { get; set; } = new List<string>();
    public string OutputMode { get; set; } = "stdout";
    public int TimeoutSeconds { get; set; } = 120;
    public int MaxOutputBytes { get; set; } = 16 * 1024 * 1024;
}

internal sealed class ReaderComparisonConfiguration {
    public List<ReaderComparisonRunnerConfiguration> Runners { get; set; } = new List<ReaderComparisonRunnerConfiguration>();
}

internal sealed class ReaderComparisonProbeResult {
    public string Id { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public bool Applied { get; set; }
    public bool Passed { get; set; }
}

internal sealed class ReaderComparisonCaseResult {
    public string CaseId { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string? Error { get; set; }
    public string MarkdownSha256 { get; set; } = string.Empty;
    public bool Deterministic { get; set; }
    public double DurationMilliseconds { get; set; }
    public long? AllocatedBytes { get; set; }
    public long? PeakWorkingSetBytes { get; set; }
    public int PassedProbes { get; set; }
    public int AppliedProbes { get; set; }
    public IReadOnlyList<ReaderComparisonProbeResult> Probes { get; set; } = Array.Empty<ReaderComparisonProbeResult>();
}

internal sealed class ReaderComparisonToolResult {
    public string Tool { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string? Error { get; set; }
    public IReadOnlyList<ReaderComparisonCaseResult> Cases { get; set; } = Array.Empty<ReaderComparisonCaseResult>();
}

internal sealed class ReaderComparisonReport {
    public int SchemaVersion { get; set; } = 1;
    public DateTimeOffset CreatedUtc { get; set; }
    public string Runtime { get; set; } = string.Empty;
    public string OperatingSystem { get; set; } = string.Empty;
    public IReadOnlyList<ReaderComparisonToolResult> Tools { get; set; } = Array.Empty<ReaderComparisonToolResult>();
}
