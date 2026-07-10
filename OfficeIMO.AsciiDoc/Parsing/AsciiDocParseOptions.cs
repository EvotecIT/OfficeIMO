namespace OfficeIMO.AsciiDoc;

/// <summary>Options controlling lossless AsciiDoc parsing.</summary>
public sealed class AsciiDocParseOptions {
    /// <summary>
    /// Maximum accepted UTF-16 source length. Defaults to 64 MiB of characters. Set to null to disable the limit.
    /// </summary>
    public int? MaximumInputLength { get; set; } = 64 * 1024 * 1024;

    /// <summary>
    /// Maximum number of top-level source blocks. Defaults to 1,000,000. Set to null to disable the limit.
    /// </summary>
    public int? MaximumBlockCount { get; set; } = 1_000_000;

    /// <summary>Maximum nested inline formatting depth. Defaults to 64.</summary>
    public int MaximumInlineNestingDepth { get; set; } = 64;

    /// <summary>Maximum inline nodes created for one document. Defaults to 1,000,000.</summary>
    public int MaximumInlineNodeCount { get; set; } = 1_000_000;
}
