namespace OfficeIMO.Bibliography;

/// <summary>Output representation for a CSL rendering operation.</summary>
public enum CslOutputFormat {
    /// <summary>Unicode text without presentation markup.</summary>
    PlainText,
    /// <summary>Escaped HTML with CSL presentation markup.</summary>
    Html
}

/// <summary>Local data and resource limits for loading a CSL style.</summary>
public sealed class CslStyleLoadOptions {
    /// <summary>Maximum characters in each style XML document.</summary>
    public int MaximumCharacters { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum XML and macro nesting depth.</summary>
    public int MaximumNestingDepth { get; set; } = 64;
    /// <summary>Resolves an independent parent style by its CSL identifier. The library performs no network access.</summary>
    public Func<string, string?>? IndependentStyleResolver { get; set; }
}

/// <summary>Local rendering configuration, copied when a processor is created.</summary>
public sealed class CslRenderOptions {
    /// <summary>Output representation. The default is plain text.</summary>
    public CslOutputFormat OutputFormat { get; set; }
    /// <summary>Links rendered bibliography URL, DOI, PMID, and PMCID values in HTML output. Defaults to true.</summary>
    /// <remarks>Only absolute HTTP and HTTPS targets are linked. Citation clusters and plain text output do not contain links.</remarks>
    public bool LinkBibliographyIdentifiers { get; set; } = true;
    /// <summary>Requested language dialect, overriding the style default.</summary>
    public string? Locale { get; set; }
    /// <summary>Additional CSL locale XML documents indexed by their language code.</summary>
    public IDictionary<string, string> Locales { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    /// <summary>Caller-provided abbreviations indexed first by CSL variable and then by the unabbreviated value.</summary>
    public IDictionary<string, IDictionary<string, string>> Abbreviations { get; } = new Dictionary<string, IDictionary<string, string>>(StringComparer.Ordinal);
    /// <summary>Maximum characters in the combined output of an operation.</summary>
    public int MaximumOutputCharacters { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum characters in either representation of an intermediate rendering or sort key.</summary>
    /// <remarks>Separate from the final output limit because escaped HTML and style markup can be longer than plain text.</remarks>
    public int MaximumIntermediateCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum citation items in a rendering operation.</summary>
    public int MaximumCitationItems { get; set; } = 100000;
    /// <summary>Maximum rendering work units per operation, including element evaluations and sort comparisons.</summary>
    /// <remarks>This bounds repeated macro expansion and sorting work. Disambiguation evaluations share the same budget.</remarks>
    public int MaximumRenderingOperations { get; set; } = 10000000;
    /// <summary>Rejects citation-data conversion when the source bibliography contains semantics the CSL data projection cannot preserve.</summary>
    public bool RequireNoDataLoss { get; set; }
}

/// <summary>A reference to a bibliography item within a citation.</summary>
public sealed class CslCitationItem {
    /// <summary>Creates a citation reference by its exact bibliography key.</summary>
    public CslCitationItem(string key) {
        Key = string.IsNullOrWhiteSpace(key) ? throw new ArgumentException("A citation key is required.", nameof(key)) : key;
    }
    /// <summary>Bibliography key.</summary>
    public string Key { get; }
    /// <summary>Locator value, such as a page range.</summary>
    public string? Locator { get; set; }
    /// <summary>Locator type. Defaults to page.</summary>
    public string LocatorType { get; set; } = "page";
    /// <summary>Literal text before this reference.</summary>
    public string? Prefix { get; set; }
    /// <summary>Literal text after this reference.</summary>
    public string? Suffix { get; set; }
    /// <summary>Omits the author when the surrounding prose names the author.</summary>
    public bool SuppressAuthor { get; set; }
    /// <summary>Renders the style's first names expression for a narrative citation, using its substitutions and formatting.</summary>
    /// <remarks>When the style renders no names, author names use default long formatting. Citation-layout affixes are omitted for a cluster containing only narrative items. Cannot be combined with <see cref="SuppressAuthor"/>.</remarks>
    public bool AuthorOnly { get; set; }
}

/// <summary>A citation cluster in document order.</summary>
public sealed class CslCitation {
    /// <summary>Creates a citation with a stable caller-owned identifier.</summary>
    public CslCitation(string id) {
        Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("A citation identifier is required.", nameof(id)) : id;
    }
    /// <summary>Stable cluster identifier.</summary>
    public string Id { get; }
    /// <summary>Note number for note styles, or zero for an in-text citation.</summary>
    public int NoteIndex { get; set; }
    /// <summary>Whether host text precedes this citation in its note. Defaults to false.</summary>
    /// <remarks>Set this for a citation inserted into existing footnote or endnote prose. It prevents automatic note-opening capitalization. Ignored for body citations and in-text styles. An item prefix containing non-whitespace text also prevents automatic capitalization.</remarks>
    public bool NoteHasPrecedingText { get; set; }
    /// <summary>References in source order, before style-specific sorting.</summary>
    public IList<CslCitationItem> Items { get; } = new List<CslCitationItem>();
}

/// <summary>A rendered citation or bibliography entry.</summary>
public sealed class CslRenderedEntry {
    internal CslRenderedEntry(string key, string content, bool isEmpty) { Key = key; Content = content; IsEmpty = isEmpty; }
    /// <summary>Citation identifier or bibliography key.</summary>
    public string Key { get; }
    /// <summary>Rendered output in the requested representation.</summary>
    public string Content { get; }
    /// <summary>Whether the style produced no text for this entry, independently of HTML entry wrappers.</summary>
    /// <remarks>Empty entries retain their keys and numbering in the operation result. Hosts may omit them from display. Whitespace explicitly rendered by the style is text and does not mark an entry empty.</remarks>
    public bool IsEmpty { get; }
}

/// <summary>An operation result with citations in document order and the style-sorted bibliography.</summary>
public sealed class CslRenderResult {
    internal CslRenderResult(IReadOnlyList<CslRenderedEntry> citations, IReadOnlyList<CslRenderedEntry> bibliography, CslBibliographyLayout? bibliographyLayout) {
        Citations = citations; Bibliography = bibliography; BibliographyLayout = bibliographyLayout;
    }
    /// <summary>Rendered citation clusters.</summary>
    public IReadOnlyList<CslRenderedEntry> Citations { get; }
    /// <summary>Rendered bibliography entries.</summary>
    public IReadOnlyList<CslRenderedEntry> Bibliography { get; }
    /// <summary>Style spacing, indentation, and alignment settings, or null when the style has no bibliography.</summary>
    public CslBibliographyLayout? BibliographyLayout { get; }
}
