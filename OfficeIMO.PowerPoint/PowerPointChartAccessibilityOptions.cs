namespace OfficeIMO.PowerPoint {
    /// <summary>Accessibility metadata applied while authoring a native chart.</summary>
    public sealed class PowerPointChartAccessibilityOptions {
        /// <summary>Concise alternative text describing the chart's purpose or conclusion.</summary>
        public string? AlternativeText { get; set; }
        /// <summary>Optional caller-authored plain-text data summary.</summary>
        public string? DataSummary { get; set; }
        /// <summary>Whether to append the data summary to the native chart alternative text.</summary>
        public bool IncludeDataSummaryInAlternativeText { get; set; } = true;
        /// <summary>Optional native shape name.</summary>
        public string? Name { get; set; }
    }
}
