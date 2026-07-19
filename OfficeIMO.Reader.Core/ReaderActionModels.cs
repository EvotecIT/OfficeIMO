namespace OfficeIMO.Reader;

/// <summary>
/// Source-neutral action locations exposed by Reader adapters.
/// </summary>
public enum ReaderActionScope {
    /// <summary>The source did not expose a recognized action scope.</summary>
    Unknown,
    /// <summary>Document-level open action.</summary>
    DocumentOpen,
    /// <summary>Catalog-level active action.</summary>
    Catalog,
    /// <summary>Page-level active action.</summary>
    Page,
    /// <summary>Annotation-level active action.</summary>
    Annotation
}

/// <summary>
/// Passive action metadata extracted from a source document.
/// This summary never carries executable payloads such as JavaScript bodies.
/// </summary>
public sealed class ReaderActionSummary {
    /// <summary>Source-neutral action scope.</summary>
    public ReaderActionScope Scope { get; set; }

    /// <summary>Source-specific action type, for example JavaScript, Launch, Destination, or GoTo.</summary>
    public string ActionType { get; set; } = string.Empty;

    /// <summary>Source container that exposed the action, when available.</summary>
    public string? Source { get; set; }

    /// <summary>Name-tree key, catalog slot, or other source action name, when available.</summary>
    public string? Name { get; set; }

    /// <summary>Additional-action trigger key, when available.</summary>
    public string? TriggerName { get; set; }

    /// <summary>Stable action path including chained actions, when available.</summary>
    public string? ActionPath { get; set; }

    /// <summary>One-based source page number associated with the action, when known.</summary>
    public int? PageNumber { get; set; }

    /// <summary>True when the action was found through a chained action path.</summary>
    public bool IsChainedAction { get; set; }

    /// <summary>True when the action type can execute script, launch external content, submit/import data, or play rich media.</summary>
    public bool IsPotentiallyUnsafe { get; set; }

    /// <summary>One-based destination page number for safe navigation actions, when known.</summary>
    public int? DestinationPageNumber { get; set; }

    /// <summary>Viewer destination mode for safe navigation actions, when known.</summary>
    public string? DestinationMode { get; set; }

    /// <summary>Destination top coordinate for safe navigation actions, when present.</summary>
    public double? DestinationTop { get; set; }

    /// <summary>Destination left coordinate for safe navigation actions, when present.</summary>
    public double? DestinationLeft { get; set; }

    /// <summary>Destination bottom coordinate for safe navigation actions, when present.</summary>
    public double? DestinationBottom { get; set; }

    /// <summary>Destination right coordinate for safe navigation actions, when present.</summary>
    public double? DestinationRight { get; set; }
}
