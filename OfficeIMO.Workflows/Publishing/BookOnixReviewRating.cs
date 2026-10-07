namespace OfficeIMO.Workflows;

/// <summary>A plain-text name for a review's rating units, optionally tagged with an ONIX list 74 language.</summary>
public sealed record BookOnixRatingUnit(string Text, string? LanguageCode = null);

/// <summary>A publisher-supplied review score. Value must be nonnegative; an optional positive integer limit
/// must be at least the score. No default scale, unit, verification or score calculation is implied.</summary>
public sealed record BookOnixReviewRating(decimal Value, int? Limit = null) {
    /// <summary>Up to 16 plain-text unit translations, at most 50 UTF-16 code units each.
    /// Repeated units require distinct explicit languages; a singleton may omit its language.</summary>
    public IReadOnlyList<BookOnixRatingUnit> Units { get; init; } = [];
}
