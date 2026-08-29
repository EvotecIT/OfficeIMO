namespace OfficeIMO.Bibliography;

/// <summary>A format-neutral citation-data record.</summary>
public sealed class BibliographyItem {
    internal IDictionary<string, string> BibFieldNames { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Stable citation key or source identifier.</summary>
    public string Key { get; set; } = string.Empty;
    /// <summary>Format-neutral item type.</summary>
    public BibliographyItemType Type { get; set; }
    /// <summary>Original native item type, when available.</summary>
    public string? NativeType { get; set; }
    /// <summary>Title.</summary>
    public string? Title { get; set; }
    /// <summary>Collection, journal, or containing work title.</summary>
    public string? ContainerTitle { get; set; }
    /// <summary>Collection title.</summary>
    public string? CollectionTitle { get; set; }
    /// <summary>Publisher.</summary>
    public string? Publisher { get; set; }
    /// <summary>Publisher place.</summary>
    public string? PublisherPlace { get; set; }
    /// <summary>Edition.</summary>
    public string? Edition { get; set; }
    /// <summary>Volume.</summary>
    public string? Volume { get; set; }
    /// <summary>Issue or number.</summary>
    public string? Issue { get; set; }
    /// <summary>Page range or article number.</summary>
    public string? Pages { get; set; }
    /// <summary>Abstract.</summary>
    public string? Abstract { get; set; }
    /// <summary>Language code or source language label.</summary>
    public string? Language { get; set; }
    /// <summary>Resource URL.</summary>
    public string? Url { get; set; }
    /// <summary>Contributors in source order.</summary>
    public IList<BibliographyContributor> Contributors { get; } = new List<BibliographyContributor>();
    /// <summary>Dates in source order.</summary>
    public IList<BibliographyDate> Dates { get; } = new List<BibliographyDate>();
    /// <summary>Typed identifiers in source order.</summary>
    public IList<BibliographyIdentifier> Identifiers { get; } = new List<BibliographyIdentifier>();
    /// <summary>Keywords in source order.</summary>
    public IList<string> Keywords { get; } = new List<string>();
    /// <summary>Notes in source order.</summary>
    public IList<string> Notes { get; } = new List<string>();
    /// <summary>Unknown or not-yet-typed native fields in source order.</summary>
    public IList<BibliographyNativeField> NativeFields { get; } = new List<BibliographyNativeField>();

    /// <summary>Gets the first identifier matching a scheme, ignoring case.</summary>
    public string? GetIdentifier(string scheme) {
        if (scheme == null) throw new ArgumentNullException(nameof(scheme));
        return Identifiers.FirstOrDefault(identifier => string.Equals(identifier.Scheme, scheme, StringComparison.OrdinalIgnoreCase))?.Value;
    }

    /// <summary>Sets or removes the first identifier matching a scheme.</summary>
    public void SetIdentifier(string scheme, string? value) {
        if (string.IsNullOrWhiteSpace(scheme)) throw new ArgumentException("Identifier scheme cannot be empty.", nameof(scheme));
        BibliographyIdentifier? existing = Identifiers.FirstOrDefault(identifier => string.Equals(identifier.Scheme, scheme, StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(value)) {
            if (existing != null) Identifiers.Remove(existing);
        } else if (existing == null) {
            Identifiers.Add(new BibliographyIdentifier(scheme, value!));
        } else {
            existing!.Value = value!.Trim();
        }
    }

    /// <summary>Gets the first date having the requested role.</summary>
    public BibliographyDate? GetDate(BibliographyDateRole role) => Dates.FirstOrDefault(date => date.Role == role);
}
