using OfficeIMO.Ocr;

namespace OfficeIMO.Ocr.Process;

/// <summary>Configures the generic JSON file-protocol OCR process engine.</summary>
public sealed class ProcessOcrEngineOptions {
    /// <summary>Direct executable path or name. Required.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>
    /// Argument templates. Supported placeholders are <c>{request}</c>, <c>{input}</c>, <c>{output}</c>,
    /// <c>{language}</c>, <c>{candidateId}</c>, <c>{sourceId}</c>, and <c>{pageNumber}</c>.
    /// </summary>
    public IReadOnlyList<string> Arguments { get; set; } = new[] { "{request}" };

    /// <summary>
    /// Stable engine identifier. Defaults to <c>process</c>; the untrimmed value must not exceed
    /// <see cref="OcrEngineRunner.MaximumEngineIdCharacters"/> characters.
    /// </summary>
    public string Id { get; set; } = "process";

    /// <summary>Optional working directory for the executable.</summary>
    public string? WorkingDirectory { get; set; }

    /// <summary>Optional environment values applied to the child process.</summary>
    public IReadOnlyDictionary<string, string> EnvironmentVariables { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>Optional parent directory for isolated per-request temporary folders.</summary>
    public string? TemporaryDirectory { get; set; }

    /// <summary>Maximum process duration. Defaults to two minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>Maximum JSON response size. Defaults to 16 MiB.</summary>
    public long MaxOutputBytes { get; set; } = 16L * 1024L * 1024L;

    /// <summary>Maximum input payload size accepted by direct engine calls. Defaults to 25 MiB.</summary>
    public long MaxInputBytes { get; set; } = 25L * 1024L * 1024L;

    /// <summary>Maximum retained standard-output and standard-error characters.</summary>
    public int MaxProcessOutputCharacters { get; set; } = 64 * 1024;

    /// <summary>Whether isolated temporary files are retained after recognition.</summary>
    public bool KeepTemporaryFiles { get; set; }

    /// <summary>Capabilities advertised by the external protocol implementation.</summary>
    public OcrEngineCapabilities Capabilities { get; set; } = new OcrEngineCapabilities {
        SupportsLineSpans = true,
        SupportsWordSpans = true,
        SupportsCharacterSpans = true,
        SupportsConfidence = true,
        SupportsConcurrentRequests = true
    };
}
