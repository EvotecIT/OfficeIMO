namespace OfficeIMO.Chm;

/// <summary>Selection and aggregate budgets shared by CHM conversion adapters.</summary>
public sealed class ChmConversionOptions {
    /// <summary>Maximum aggregate parsed nodes across selected topics and nodes retained in a combined HTML book.</summary>
    public int MaxHtmlNodes { get; set; } = 1_000_000;
    /// <summary>Optional topic paths. Selected topics retain book order; unknown or non-topic paths are rejected.</summary>
    public IReadOnlyList<string>? TopicPaths { get; set; }
    /// <summary>Maximum topics in one export. The operation fails rather than truncating.</summary>
    public int MaxTopics { get; set; } = 20_000;
    /// <summary>Maximum decoded source HTML characters across selected topics.</summary>
    public int MaxTotalHtmlCharacters { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum serialized output bytes.</summary>
    public long MaxOutputBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Embeds archive images as data URLs when projecting a combined HTML book or Markdown. PDF and EPUB use archive resolvers.</summary>
    public bool EmbedImages { get; set; } = true;
    /// <summary>Maximum image bytes per embedded image.</summary>
    public int MaxEmbeddedImageBytes { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum image bytes embedded in one HTML/Markdown projection, counting repeated references.</summary>
    public long MaxTotalEmbeddedImageBytes { get; set; } = 32L * 1024 * 1024;
    /// <summary>Creates a detached options snapshot.</summary>
    public ChmConversionOptions Clone() {
        var copy = (ChmConversionOptions)MemberwiseClone();
        copy.TopicPaths = TopicPaths == null ? null : Array.AsReadOnly(TopicPaths.ToArray());
        return copy;
    }
    internal void Validate() {
        if (MaxHtmlNodes < 1) throw new ArgumentOutOfRangeException(nameof(MaxHtmlNodes));
        if (MaxTopics < 1) throw new ArgumentOutOfRangeException(nameof(MaxTopics));
        if (MaxTotalHtmlCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxTotalHtmlCharacters));
        if (MaxOutputBytes < 1 || MaxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxOutputBytes));
        if (MaxEmbeddedImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxEmbeddedImageBytes));
        if (MaxTotalEmbeddedImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxTotalEmbeddedImageBytes));
    }
}
