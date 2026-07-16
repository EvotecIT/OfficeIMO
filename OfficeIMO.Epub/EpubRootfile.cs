namespace OfficeIMO.Epub;

/// <summary>Represents a rootfile declaration from META-INF/container.xml.</summary>
public sealed class EpubRootfile {
    /// <summary>Normalized package path in the EPUB container.</summary>
    public string FullPath { get; internal set; } = string.Empty;

    /// <summary>Declared rootfile media type.</summary>
    public string? MediaType { get; internal set; }

    /// <summary>Whether the declared rootfile exists in the EPUB container.</summary>
    public bool IsAvailable { get; internal set; }

    /// <summary>Whether this rootfile was selected for extraction.</summary>
    public bool IsSelected { get; internal set; }
}
