namespace OfficeIMO.Epub;

/// <summary>A changed text block, matched by explicit ID or structural position within a resource.</summary>
public sealed class EpubEditionTextChange {
    internal EpubEditionTextChange(string manifestId, string locator, string? before, string? after) {
        ManifestId = manifestId; Locator = locator;
        IsTruncated = before?.Length > 4096 || after?.Length > 4096;
        PreviousText = Clip(before); CurrentText = Clip(after);
    }
    /// <summary>Manifest identity of the text resource.</summary>
    public string ManifestId { get; }
    /// <summary>Explicit #id when unique, otherwise a namespace-aware element-index path.</summary>
    public string Locator { get; }
    /// <summary>Original text, capped at 4096 UTF-16 characters, or null for an added block.</summary>
    public string? PreviousText { get; }
    /// <summary>Revised text, capped at 4096 UTF-16 characters, or null for a removed block.</summary>
    public string? CurrentText { get; }
    /// <summary>Whether either text excerpt was clipped. Change detection always compares complete text.</summary>
    public bool IsTruncated { get; }
    private static string? Clip(string? text) {
        if (text == null || text.Length <= 4096) return text;
        return text.Substring(0, char.IsHighSurrogate(text[4095]) ? 4095 : 4096);
    }
}
