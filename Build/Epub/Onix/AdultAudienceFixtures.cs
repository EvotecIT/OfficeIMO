using OfficeIMO.Workflows;

internal static class AdultAudienceFixtures {
    internal static BookOnixAudienceMetadata? Create(string profile) => profile switch {
        "adult-unrated" => new() { Categories = [new(BookOnixAudienceType.GeneralAdult, true)],
            AdultRatings = [new(BookOnixAdultAudienceRating.Unrated)] },
        "adult-general" => new() { Categories = [new(BookOnixAudienceType.GeneralAdult, true)],
            AdultRatings = [new(BookOnixAdultAudienceRating.AnyAdultAudience)] },
        "adult-advice" => new() { Categories = [new(BookOnixAudienceType.GeneralAdult, true)],
            AdultRatings = Enum.GetValues<BookOnixAdultAudienceRating>().Where(rating => rating >= BookOnixAdultAudienceRating.ContentAdvice)
                .Select(rating => new BookOnixAdultAudience(rating, rating == BookOnixAdultAudienceRating.ContentAdvice) {
                    Headings = rating == BookOnixAdultAudienceRating.ContentAdvice
                        ? [new("Synthetic publisher advice & context", "eng"), new("Przykładowa informacja wydawcy", "pol")] : []
                }).ToArray() },
        _ => null
    };
}
