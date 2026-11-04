namespace PSWriteOffice.Models.Confluence;

/// <summary>Projection returned for each Confluence page.</summary>
public enum OfficeConfluencePageOutput {
    /// <summary>The native page.</summary>
    Page,
    /// <summary>Markdown and native fidelity evidence.</summary>
    Markdown,
    /// <summary>HTML and native fidelity evidence.</summary>
    Html
}
