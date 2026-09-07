namespace OfficeIMO.AI;

/// <summary>Hard bounds for a document operation. All batches share the request and duration budgets.</summary>
public sealed record OfficeAiLimits {
    /// <summary>Maximum source bytes copied from a file or caller stream.</summary>
    public int MaxInputBytes { get; init; } = 32 * 1024 * 1024;
    /// <summary>Maximum extracted text retained in a snapshot.</summary>
    public int MaxDocumentCharacters { get; init; } = 2_000_000;
    /// <summary>Maximum evidence records retained in a snapshot.</summary>
    public int MaxDocumentBlocks { get; init; } = 50_000;
    /// <summary>Maximum page count in a snapshot.</summary>
    public int MaxPages { get; init; } = 500;
    /// <summary>Maximum serialized evidence and user instructions per model request.</summary>
    public int MaxRequestCharacters { get; init; } = 48_000;
    /// <summary>Maximum model response characters per batch.</summary>
    public int MaxResponseCharacters { get; init; } = 64_000;
    /// <summary>Maximum model requests in the whole operation; no automatic repair requests are made.</summary>
    public int MaxRequests { get; init; } = 32;
    /// <summary>Maximum hierarchical summary reduction passes sharing the operation request budget.</summary>
    public int MaxSynthesisPasses { get; init; } = 3;
    /// <summary>Maximum records of each result kind per batch.</summary>
    public int MaxResultItems { get; init; } = 200;
    /// <summary>Maximum cells in a returned table.</summary>
    public int MaxTableCells { get; init; } = 10_000;
    /// <summary>Maximum aggregate encoded image bytes per request.</summary>
    public int MaxImageBytes { get; init; } = 8 * 1024 * 1024;
    /// <summary>Maximum aggregate image pixels per request.</summary>
    public long MaxImagePixels { get; init; } = 20_000_000;
    /// <summary>Total wall-clock deadline for inference and validation across every batch.</summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromMinutes(3);

    internal void Validate() {
        if (MaxInputBytes is < 1 or > 268_435_456 || MaxDocumentCharacters is < 1 or > 20_000_000
            || MaxDocumentBlocks is < 1 or > 200_000 || MaxPages is < 1 or > 10_000
            || MaxRequestCharacters is < 4096 or > 2_000_000 || MaxResponseCharacters is < 1024 or > 2_000_000
            || MaxSynthesisPasses is < 1 or > 8 || MaxRequests is < 1 or > 256 || MaxResultItems is < 1 or > 200
            || MaxTableCells is < 1 or > 100_000 || MaxImageBytes is < 1 or > 67_108_864
            || MaxImagePixels is < 1 or > 100_000_000 || Timeout <= TimeSpan.Zero || Timeout > TimeSpan.FromHours(1))
            throw new ArgumentOutOfRangeException(nameof(OfficeAiLimits), "Document AI limits are outside the supported bounds.");
    }
}
