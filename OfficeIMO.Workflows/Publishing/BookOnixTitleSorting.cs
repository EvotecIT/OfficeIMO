namespace OfficeIMO.Workflows;

/// <summary>
/// Explicit alphabetical-sorting assertion for a title. An instance with no Prefix asserts NoPrefix.
/// Omit the containing TitleSorting property when the sorting prefix is unknown.
/// </summary>
public sealed record BookOnixTitleSorting {
    /// <summary>
    /// Exact leading text excluded from sorting, including any separator to remove, such as "The ".
    /// Matching is ordinal and case-sensitive; the remaining title must be nonblank.
    /// Null explicitly asserts that the full title is used for sorting. Empty or whitespace-only values are invalid.
    /// </summary>
    public string? Prefix { get; init; }
}
