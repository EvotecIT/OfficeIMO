namespace OfficeIMO.Markdown;

/// <summary>
/// Placement anchors for block parser extensions within the default markdown reader pipeline.
/// </summary>
public enum MarkdownBlockParserPlacement {
    /// <summary>After front matter parsing and before quote parsing.</summary>
    AfterFrontMatter,
    /// <summary>After HTML block parsing and before reference-link definitions.</summary>
    AfterHtmlBlocks,
    /// <summary>After reference-link definitions and before table/list parsing.</summary>
    AfterReferenceLinkDefinitions,
    /// <summary>Immediately before paragraph parsing (the pipeline fallback).</summary>
    BeforeParagraphs
}
