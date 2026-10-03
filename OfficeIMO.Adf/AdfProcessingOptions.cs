using System.Threading;

namespace OfficeIMO.Adf;

/// <summary>Resource limits for ADF parsing, writing, validation and projections.</summary>
/// <remarks>Limits reject the operation rather than truncating content. Cancellation is checked at ADF traversal boundaries; synchronous Markdown/HTML parsers are checked before and after their calls.</remarks>
public class AdfProcessingOptions {
    /// <summary>Maximum UTF-8 input size for JSON, Markdown or HTML text. Defaults to 16 MiB.</summary>
    public int MaxInputBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum content-node nesting depth, from 1 through 256. Defaults to 64. Markdown object input is checked against this depth before conversion.</summary>
    public int MaxDepth { get; set; } = 64;
    /// <summary>Maximum number of nodes and marks visited, including repeated references. Defaults to 100,000. Markdown object input is also checked against this object count.</summary>
    public int MaxNodes { get; set; } = 100_000;
    /// <summary>Maximum aggregate text in nodes and JSON attribute/extension strings, in UTF-16 characters. Defaults to 16 Mi characters.</summary>
    public int MaxTextCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum serialized JSON size in UTF-8 bytes. Defaults to 32 MiB.</summary>
    public int MaxOutputBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum Markdown or HTML result length in UTF-16 characters. Defaults to 32 Mi characters.</summary>
    public int MaxOutputCharacters { get; set; } = 32 * 1024 * 1024;
    /// <summary>Cancellation observed during ADF traversal and JSON writes.</summary>
    public CancellationToken CancellationToken { get; set; }

    internal AdfProcessingOptions WithCancellation(CancellationToken token) => new AdfProcessingOptions {
        MaxInputBytes = MaxInputBytes, MaxDepth = MaxDepth, MaxNodes = MaxNodes,
        MaxTextCharacters = MaxTextCharacters, MaxOutputBytes = MaxOutputBytes,
        MaxOutputCharacters = MaxOutputCharacters, CancellationToken = token
    };

    internal void Check() {
        if (MaxInputBytes <= 0 || MaxDepth <= 0 || MaxDepth > 256 || MaxNodes <= 0 || MaxTextCharacters <= 0 || MaxOutputBytes <= 0 || MaxOutputCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(AdfProcessingOptions), "ADF limits must be positive, and MaxDepth must not exceed 256.");
        CancellationToken.ThrowIfCancellationRequested();
    }
}
