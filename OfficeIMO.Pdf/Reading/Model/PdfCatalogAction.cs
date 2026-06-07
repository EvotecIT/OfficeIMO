namespace OfficeIMO.Pdf;

/// <summary>
/// A catalog-level active action discovered from a supported PDF catalog action slot.
/// </summary>
public sealed class PdfCatalogAction {
    internal PdfCatalogAction(string name, string actionType, string source, string? triggerName = null) {
        Name = name;
        ActionType = actionType;
        Source = source;
        TriggerName = triggerName;
    }

    /// <summary>Name-tree key or catalog action slot associated with the action.</summary>
    public string Name { get; }

    /// <summary>PDF action type, for example JavaScript.</summary>
    public string ActionType { get; }

    /// <summary>Catalog source that contained the action, for example Names/JavaScript, OpenAction, or AA.</summary>
    public string Source { get; }

    /// <summary>Catalog additional-action trigger name when the action came from /AA.</summary>
    public string? TriggerName { get; }
}
