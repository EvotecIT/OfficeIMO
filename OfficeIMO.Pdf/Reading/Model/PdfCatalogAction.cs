namespace OfficeIMO.Pdf;

/// <summary>
/// A catalog-level active action discovered from a supported PDF catalog action slot.
/// </summary>
public sealed class PdfCatalogAction {
    internal PdfCatalogAction(string name, string actionType, string source, string? triggerName = null, string? actionPath = null, bool isChainedAction = false) {
        Name = name;
        ActionType = actionType;
        Source = source;
        TriggerName = triggerName;
        ActionPath = actionPath;
        IsChainedAction = isChainedAction;
    }

    /// <summary>Name-tree key or catalog action slot associated with the action.</summary>
    public string Name { get; }

    /// <summary>PDF action type, for example JavaScript.</summary>
    public string ActionType { get; }

    /// <summary>Catalog source that contained the action, for example Names/JavaScript, OpenAction, or AA.</summary>
    public string Source { get; }

    /// <summary>Catalog additional-action trigger name when the action came from /AA.</summary>
    public string? TriggerName { get; }

    /// <summary>Stable catalog action path including chained /Next actions, when available.</summary>
    public string? ActionPath { get; }

    /// <summary>True when this catalog action was discovered through a chained /Next path.</summary>
    public bool IsChainedAction { get; }
}
