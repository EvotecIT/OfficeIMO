namespace OfficeIMO.Epub;

/// <summary>Standard EPUB 3 title classifications.</summary>
public enum EpubTitleKind {
    /// <summary>The principal title.</summary>
    Main,
    /// <summary>A subtitle.</summary>
    Subtitle,
    /// <summary>An abbreviated title.</summary>
    Short,
    /// <summary>A title shared with a collection.</summary>
    Collection,
    /// <summary>An edition statement.</summary>
    Edition,
    /// <summary>A title combining multiple title components.</summary>
    Expanded
}

/// <summary>A title with its EPUB 3 classification, display order and sorting form.</summary>
public sealed class EpubTitleMetadata {
    /// <summary>Title text.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Standard title classification.</summary>
    public EpubTitleKind Kind { get; set; } = EpubTitleKind.Main;
    /// <summary>Optional relative display sequence. Does not reorder XML declarations.</summary>
    public uint? DisplaySequence { get; set; }
    /// <summary>Optional sorting form.</summary>
    public string? FileAs { get; set; }
    /// <summary>Optional BCP 47 language of the title.</summary>
    public string? Language { get; set; }
}
