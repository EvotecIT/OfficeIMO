namespace OfficeIMO.Epub;

/// <summary>Updates to primary descriptive Dublin Core values. Null fields retain existing values.</summary>
public sealed class EpubPublicationDetails {
    /// <summary>Publisher name.</summary>
    public string? Publisher { get; set; }
    /// <summary>Plain-text description.</summary>
    public string? Description { get; set; }
    /// <summary>Rights statement supplied by the publisher.</summary>
    public string? Rights { get; set; }
    /// <summary>Calendar publication date. Only the date component is used; no timezone conversion occurs.</summary>
    public DateTime? PublicationDate { get; set; }
}
