using OfficeIMO.Workflows;

internal static class ReviewRatingFixtures {
    internal static IReadOnlyList<BookOnixCollateralText> Create() => [
        new() { Type = BookOnixTextType.ReviewQuote, Audiences = [BookOnixContentAudience.EndCustomers],
            Texts = [new("<p>A synthetic <strong>review</strong> quotation.</p>", "eng") { Format = BookOnixCollateralTextFormat.Xhtml }],
            ReviewRating = new(4.50m, 5) { Units = [new("stars & points", "eng"), new("gwiazdki", "pol")] },
            Authors = ["Example Reviewer"], SourceTitles = [new("Example Journal", "eng")] },
        new() { Type = BookOnixTextType.PreviousEditionReview, Audiences = [BookOnixContentAudience.Unrestricted],
            Texts = [new("Synthetic previous-edition review.")], ReviewRating = new(0) },
        new() { Type = BookOnixTextType.PreviousWorkReview, Audiences = [BookOnixContentAudience.BookTrade],
            Texts = [new("Synthetic previous-work review.")], ReviewRating = new(100, 100) { Units = [new("points")] } }
    ];
}
