using OfficeIMO.Workflows;

internal static class CollectionFixtures {
    internal static IReadOnlyList<BookOnixCollection> Create() => [
        new() { Type = BookOnixCollectionType.Publisher, Title = "Example & studies", Subtitle = "Collected works", LanguageCode = "eng",
            Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "set-1", "Example catalog"),
                new(BookOnixCollectionIdentifierType.Issn, "0317-8471"), new(BookOnixCollectionIdentifierType.Isbn13, "9781861972712")],
            Sequences = Enum.GetValues<BookOnixCollectionSequenceType>().Select(type =>
                new BookOnixCollectionSequence(type, type == BookOnixCollectionSequenceType.Proprietary ? "3.-.8" : "2.1",
                    type == BookOnixCollectionSequenceType.Proprietary ? "Curriculum order" : null)).ToArray() },
        new() { Type = BookOnixCollectionType.Editorial, Title = "Classics" },
        new() { Type = BookOnixCollectionType.Ascribed, Title = "Library selection", SourceName = "Example Library" }
    ];
}
