namespace OfficeIMO.Pdf;

/// <summary>
/// Metadata for an action attached to a page dictionary through the page /AA additional-actions entry.
/// </summary>
public sealed class PdfPageAction {
    internal PdfPageAction(int? pageNumber, string triggerName, string actionType, string? actionPath = null) {
        PageNumber = pageNumber;
        TriggerName = triggerName;
        ActionType = actionType;
        ActionPath = string.IsNullOrEmpty(actionPath) ? triggerName : actionPath!;
    }

    /// <summary>One-based source page number when the page context is known.</summary>
    public int? PageNumber { get; }

    /// <summary>Additional-action trigger key from the page /AA dictionary, for example O or C.</summary>
    public string TriggerName { get; }

    /// <summary>PDF action type name from the action dictionary /S entry.</summary>
    public string ActionType { get; }

    /// <summary>Stable page-action path including chained /Next actions, for example O, O.Next, or O.Next.0.</summary>
    public string ActionPath { get; }

    /// <summary>True when this action was discovered through a chained /Next action.</summary>
    public bool IsChainedAction => !string.Equals(ActionPath, TriggerName, StringComparison.Ordinal);

    internal PdfPageAction WithPageNumber(int pageNumber) {
        return PageNumber == pageNumber ? this : new PdfPageAction(pageNumber, TriggerName, ActionType, ActionPath);
    }
}
