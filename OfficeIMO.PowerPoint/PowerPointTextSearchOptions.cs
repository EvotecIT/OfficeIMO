namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Comparison settings and content scope for <see cref="PowerPointDocument.ReplaceAll"/> and <see cref="PowerPointSlide.ReplaceAll"/>.
    /// </summary>
    /// <remarks>
    /// Matching follows <see cref="OfficeTextSearchOptions"/>: ordinal and case-insensitive unless
    /// <see cref="OfficeTextSearchOptions.MatchCase"/> is set.
    /// </remarks>
    public sealed class PowerPointTextSearchOptions : OfficeTextSearchOptions {
        /// <summary>Includes table cell text. The default is <c>true</c>.</summary>
        public bool IncludeTables { get; set; } = true;

        /// <summary>Includes speaker notes. The default is <c>false</c>.</summary>
        public bool IncludeNotes { get; set; }
    }
}
