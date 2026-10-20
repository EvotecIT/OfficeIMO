namespace SyntheticReportProducer;

internal sealed class CorpusSpec {
    public required string Classification { get; init; }
    public required string SourceBinding { get; init; }
    public required FontSpec[] Fonts { get; init; }
    public required VariantSpec[] Variants { get; init; }
    public required ReportSpec[] Reports { get; init; }
}

internal sealed class FontSpec {
    public required string Role { get; init; }
    public required string FileName { get; init; }
    public required string Sha256 { get; init; }
}

internal sealed class VariantSpec {
    public required string Key { get; init; }
    public required string Orientation { get; init; }
    public required string ScreenWidth { get; init; }
}

internal sealed class ReportSpec {
    public required string Key { get; init; }
    public required string Title { get; init; }
    public required MarkerSpec Intro { get; init; }
    public required string[] NavigationLinkIds { get; init; }
    public required RegionSpec[] Regions { get; init; }
    public required TableSpec[] Tables { get; init; }
    public required LinkSpec[] Links { get; init; }
    public DetailsSectionSpec? DetailsSection { get; init; }
    public required MarkerSpec Summary { get; init; }
    public required string SummaryLinkId { get; init; }
}

internal sealed class RegionSpec {
    public required string Id { get; init; }
    public required string Label { get; init; }
    public required string Title { get; init; }
    public required string Layout { get; init; }
    public required string[] TableKeys { get; init; }
    public required MarkerSpec After { get; init; }
}

internal sealed class TableSpec {
    public required string Key { get; init; }
    public required string Component { get; init; }
    public required string SurfaceId { get; init; }
    public required string PanelTitle { get; init; }
    public required string AccessibleLabel { get; init; }
    public string? Caption { get; init; }
    public required string[] Headers { get; init; }
    public required string[] FooterLabels { get; init; }
    public string? TerminalSummary { get; init; }
    public required string PanelLinkId { get; init; }
    public required RowSpec[] Rows { get; init; }
    public required MarkerSpec After { get; init; }
}

internal sealed class RowSpec {
    public required string Id { get; init; }
    public required string Label { get; init; }
    public string? DomId { get; init; }
    public required string Detail { get; init; }
    public required string Value { get; init; }
    public string? LinkId { get; init; }
    public string? Affected { get; init; }
    public string? Severity { get; init; }
    public string? Status { get; init; }
    public string? Owner { get; init; }
    public string? FirstSeen { get; init; }
    public string? LastSeen { get; init; }
    public int Occurrences { get; init; }
}

internal sealed class MarkerSpec {
    public required string Id { get; init; }
    public required string Label { get; init; }
    public required string Text { get; init; }
    public string? LinkId { get; init; }
}

internal sealed class LinkSpec {
    public required string Id { get; init; }
    public required string Label { get; init; }
    public required string Href { get; init; }
}

internal sealed class DetailsSectionSpec {
    public required string Id { get; init; }
    public required string Label { get; init; }
    public required string Text { get; init; }
    public required DetailsSpec[] Cases { get; init; }
}

internal sealed class DetailsSpec {
    public required string Id { get; init; }
    public required string SummaryLabel { get; init; }
    public required MarkerSpec Body { get; init; }
    public bool Open { get; init; }
}
