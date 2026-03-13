namespace OfficeIMO.Markdown;

/// <summary>
/// Well-known semantic fenced-block kinds shared across markdown readers, writers, and renderers.
/// Hosts may define additional values for their own contracts.
/// </summary>
public static class MarkdownSemanticKinds {
    /// <summary>Fallback semantic kind for blocks without a more specific contract.</summary>
    public const string Custom = "custom";

    /// <summary>Semantic fenced block for chart payloads.</summary>
    public const string Chart = "chart";

    /// <summary>Semantic fenced block for graph/network payloads.</summary>
    public const string Network = "network";

    /// <summary>Semantic fenced block for static data view/table payloads.</summary>
    public const string DataView = "dataview";

    /// <summary>Semantic fenced block for Mermaid diagrams.</summary>
    public const string Mermaid = "mermaid";

    /// <summary>Semantic fenced block for display math payloads.</summary>
    public const string Math = "math";
}
