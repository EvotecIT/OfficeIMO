using OfficeIMO.Workflows;

internal static class CollateralFixtures {
    internal static IReadOnlyList<BookOnixCollateralText> Create(string profile) {
        if (profile == "collateral-unicode") return [new() { Type = BookOnixTextType.ShortDescription,
            Audiences = [BookOnixContentAudience.Unrestricted],
            Texts = [new(string.Concat(Enumerable.Repeat("😀", 350)), "eng")] }];
        if (profile != "collateral") return [];
        return Enum.GetValues<BookOnixTextType>().Select(type => new BookOnixCollateralText {
            Type = type,
            Audiences = type == BookOnixTextType.Description
                ? Enum.GetValues<BookOnixContentAudience>().Where(a => a != BookOnixContentAudience.Unrestricted).ToArray()
                : [BookOnixContentAudience.Unrestricted],
            Texts = [new(type == BookOnixTextType.Description ? new string('A', 8192) : "Synthetic <plain text> & " + type, "eng"),
                     new("Przykładowy tekst", "pol")],
            Territory = new() { Worldwide = true, ExcludedCountries = ["FR"] }, Authors = ["Example Writer"],
            SourceCorporate = "Example Review Journal", SourceTitles = [new("Example source", "eng")],
            SourceLinks = ["https://example.org/review?a=1&b=2"], PublishedOn = new(2026, 9, 1),
            UsableFrom = new(2026, 10, 1), UsableUntil = new(2027, 10, 1), UpdatedOn = new(2026, 9, 2)
        }).ToArray();
    }
}
