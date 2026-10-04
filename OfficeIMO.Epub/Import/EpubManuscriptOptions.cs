using OfficeIMO.Html;

namespace OfficeIMO.Epub;

/// <summary>Metadata, chapter boundaries and resource policy for a reflowable manuscript import.</summary>
public sealed class EpubManuscriptOptions {
    /// <summary>Overrides the source document title. A title is required when the source has none.</summary>
    public string? Title { get; set; }
    /// <summary>Overrides the source document language; otherwise its language or English is used.</summary>
    public string? Language { get; set; }
    /// <summary>Optional stable publication identity.</summary>
    public string? Identifier { get; set; }
    /// <summary>Optional primary creator.</summary>
    public string? Creator { get; set; }
    /// <summary>Starts a chapter at headings of this level or higher. Zero keeps one chapter.</summary>
    public int ChapterHeadingLevel { get; set; } = 1;
    /// <summary>Bounds retained package content independently of HTML input limits.</summary>
    public EpubPublicationLoadOptions RetentionLimits { get; set; } = new EpubPublicationLoadOptions();
    /// <summary>Adds a small responsive stylesheet for reflowable images, tables and prose.</summary>
    public bool IncludeDefaultStyles { get; set; } = true;
    /// <summary>Optional application resolver. Import never opens files or contacts the network implicitly.</summary>
    public HtmlRenderResourceResolver? ResourceResolver { get; set; }
    /// <summary>Maximum encoded bytes accepted for one manuscript resource.</summary>
    public long MaxResourceBytes { get; set; } = 10L * 1024 * 1024;
    /// <summary>Maximum encoded resource bytes accepted for one manuscript.</summary>
    public long MaxTotalResourceBytes { get; set; } = 50L * 1024 * 1024;
    /// <summary>Maximum distinct resources accepted for one manuscript.</summary>
    public int MaxResourceCount { get; set; } = 256;
    /// <summary>Creates an independent configuration snapshot for one publishing operation.</summary>
    public EpubManuscriptOptions Clone() => new EpubManuscriptOptions {
        Title = Title, Language = Language, Identifier = Identifier, Creator = Creator, ChapterHeadingLevel = ChapterHeadingLevel,
        IncludeDefaultStyles = IncludeDefaultStyles, ResourceResolver = ResourceResolver, MaxResourceBytes = MaxResourceBytes,
        MaxTotalResourceBytes = MaxTotalResourceBytes, MaxResourceCount = MaxResourceCount,
        RetentionLimits = RetentionLimits == null ? throw new InvalidOperationException("Retention limits are required.") : new EpubPublicationLoadOptions {
            MaxInputBytes = RetentionLimits.MaxInputBytes, MaxExpandedBytes = RetentionLimits.MaxExpandedBytes,
            MaxEntryBytes = RetentionLimits.MaxEntryBytes, MaxMetadataBytes = RetentionLimits.MaxMetadataBytes, MaxEntries = RetentionLimits.MaxEntries
        }
    };
}
