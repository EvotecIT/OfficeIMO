namespace OfficeIMO.Notebook.Markdown;

/// <summary>A native chart preview shared by the static and optional interactive exporters.</summary>
internal sealed record NotebookChartPreview(string ImageUri, double Width, double Height, string? Title, string Producer,
    IReadOnlyList<NotebookChartMark>? Marks) {
    internal string Label => string.IsNullOrWhiteSpace(Title) ? Producer + " chart" : Title!;
}

/// <summary>A rectangular hit region in drawing coordinates and its saved data label.</summary>
internal readonly record struct NotebookChartMark(double X, double Y, double Width, double Height, string Label);
