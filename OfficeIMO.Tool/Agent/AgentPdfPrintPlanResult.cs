namespace OfficeIMO.Tool.Agent;

/// <summary>Bounded geometry summary of a print plan; no printer queue is contacted.</summary>
public sealed class AgentPdfPrintPlanResult {
    public int SourcePageCount { get; set; }
    public int SelectedPageCount { get; set; }
    public int SheetCount { get; set; }
    public int ClippedPlacementCount { get; set; }
    /// <summary>Requested prepared-sheet color treatment; this tool does not render or deliver sheets.</summary>
    public string ColorMode { get; set; } = "color";
    public IReadOnlyList<int> SelectedPages { get; set; } = Array.Empty<int>();
    public bool Truncated { get; set; }
}
