using System.Collections.ObjectModel;

namespace OfficeIMO.AI;

/// <summary>Caller-supplied model boundary. Implementations must not grant source content tool or filesystem authority.</summary>
public interface IOfficeAiExecutor {
    /// <summary>Immutable profile describing this executor's actual configured route.</summary>
    OfficeAiExecutionProfile Profile { get; }
    /// <summary>Measures request text including transport prompt wrappers, excluding encoded image bytes. Must not execute or send evidence.</summary>
    int MeasureRequestCharacters(OfficeAiExecutionRequest request) => checked(request.Instructions.Length + request.InputJson.Length + request.OutputSchema.Length);
    /// <summary>Runs exactly one bounded request. Output is untrusted until the document engine validates it.</summary>
    Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default);
}

/// <summary>Declared route and capabilities. Declaring a capability is not a model-quality certification.</summary>
public sealed record OfficeAiExecutionProfile {
    /// <summary>Stable, non-secret profile identifier.</summary>
    public required string Id { get; init; }
    /// <summary>Provider name; do not include credentials or endpoint query strings.</summary>
    public required string Provider { get; init; }
    /// <summary>Requested model identifier.</summary>
    public required string Model { get; init; }
    /// <summary>True only for a route explicitly configured for local inference with no hosted fallback.</summary>
    public bool IsLocal { get; init; }
    /// <summary>Whether this model and transport accept inline image evidence.</summary>
    public bool SupportsImages { get; init; }
    /// <summary>Whether the provider enforces the supplied JSON schema during generation.</summary>
    public bool EnforcesJsonSchema { get; init; }
    /// <summary>Maximum request text supported by the caller's verified profile.</summary>
    public int MaxRequestCharacters { get; init; } = 48_000;
    /// <summary>Maximum aggregate inline image payload per request.</summary>
    public int MaxImageBytes { get; init; } = 8 * 1024 * 1024;

    /// <summary>Checks profile identity and resource limits before connecting or processing documents.</summary>
    public void Validate() {
        foreach (string value in new[] { Id, Provider, Model })
            if (string.IsNullOrWhiteSpace(value) || value.Length > 256 || value.Any(char.IsControl))
                throw new ArgumentException("Profile identifiers must contain 1-256 non-control characters.");
        if (MaxRequestCharacters < 4096 || MaxRequestCharacters > 2_000_000 || MaxImageBytes < 1 || MaxImageBytes > 67_108_864)
            throw new ArgumentOutOfRangeException(nameof(OfficeAiExecutionProfile));
    }
}

/// <summary>A source-bound inline image; payload access returns a defensive copy.</summary>
public sealed class OfficeAiImage {
    private readonly byte[] _bytes;
    /// <summary>Creates an image whose dimensions must have been verified by the image-producing owner.</summary>
    public OfficeAiImage(string id, int page, string mediaType, byte[] bytes, int width, int height) {
        ArgumentNullException.ThrowIfNull(bytes);
        if (string.IsNullOrWhiteSpace(id) || id.Length > 128 || id.Any(char.IsControl)) throw new ArgumentException("Invalid image identifier.", nameof(id));
        if (page < 1 || width < 1 || height < 1 || (long)width * height > 100_000_000) throw new ArgumentOutOfRangeException(nameof(width));
        if (mediaType is not ("image/png" or "image/jpeg" or "image/webp")) throw new NotSupportedException("AI image evidence accepts PNG, JPEG and WebP.");
        if (bytes.Length is < 1 or > 67_108_864) throw new ArgumentOutOfRangeException(nameof(bytes));
        Id = id; Page = page; MediaType = mediaType; Width = width; Height = height; _bytes = (byte[])bytes.Clone();
    }
    /// <summary>Evidence identifier within its source snapshot.</summary>
    public string Id { get; }
    /// <summary>One-based source page.</summary>
    public int Page { get; }
    /// <summary>Encoded media type.</summary>
    public string MediaType { get; }
    /// <summary>Decoded width supplied by the image owner.</summary>
    public int Width { get; }
    /// <summary>Decoded height supplied by the image owner.</summary>
    public int Height { get; }
    /// <summary>Encoded payload length.</summary>
    public int ByteLength => _bytes.Length;
    /// <summary>Copies the payload for a transport without exposing mutable snapshot state.</summary>
    public byte[] CopyBytes() => (byte[])_bytes.Clone();
}

/// <summary>Immutable single-batch request with no executable path or tool definitions.</summary>
public sealed record OfficeAiExecutionRequest(
    string RequestId, string Instructions, string InputJson, string OutputSchema,
    IReadOnlyList<OfficeAiImage> Images, int MaxResponseCharacters);

/// <summary>Provider response; usage may be unavailable and no monetary value is inferred.</summary>
public sealed record OfficeAiExecutionResponse(string Json, string? ResponseId = null, long? InputTokens = null,
    long? OutputTokens = null, bool IsComplete = true);

/// <summary>Observable engine stages, never hidden model reasoning.</summary>
public sealed record OfficeAiProgress(string Stage, int CompletedBatches, int TotalBatches);
