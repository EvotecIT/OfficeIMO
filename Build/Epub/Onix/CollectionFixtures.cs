using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class CollectionFixtures {
    internal static IReadOnlyList<BookOnixCollection> Create() => [
        new() { Type = BookOnixCollectionType.Publisher, Title = "Example & studies", Subtitle = "Collected works", LanguageCode = "eng",
            Contributors = [new("Series Editor", BookOnixContributorRole.SeriesEditor),
                new("Editorial & Co", BookOnixContributorRole.Editor, true),
                new("Author", BookOnixContributorRole.Author), new("Translator", BookOnixContributorRole.Translator),
                new("Illustrator", BookOnixContributorRole.Illustrator), new("Other", BookOnixContributorRole.Other)],
            Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "set-1", "Example catalog"),
                new(BookOnixCollectionIdentifierType.Issn, "0317-8471"), new(BookOnixCollectionIdentifierType.Isbn13, "9781861972712")],
            Sequences = Enum.GetValues<BookOnixCollectionSequenceType>().Select(type =>
                new BookOnixCollectionSequence(type, type == BookOnixCollectionSequenceType.Proprietary ? "3.-.8" : "2.1",
                    type == BookOnixCollectionSequenceType.Proprietary ? "Curriculum order" : null)).ToArray() },
        new() { Type = BookOnixCollectionType.Editorial, Title = "Classics", NoContributors = true },
        new() { Type = BookOnixCollectionType.Ascribed, Title = "Library selection", SourceName = "Example Library" }
    ];

    internal static void Verify(BookOnixExportResult result, XmlSchemaSet schemas) {
        XNamespace ns = BookProject.OnixNamespace;
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        var collections = document.Descendants(ns + "Collection").ToArray();
        if (collections.Length != 3 ||
            !collections[0].Elements(ns + "Contributor").Select(e => e.Element(ns + "ContributorRole")!.Value)
                .SequenceEqual(new[] { "B09", "B01", "A01", "B06", "A12", "Z99" }) ||
            collections[1].Element(ns + "NoContributor") == null || collections[2].Element(ns + "NoContributor") != null ||
            collections[2].Elements(ns + "Contributor").Any())
            throw new InvalidDataException("Collection credit assertions were not preserved.");
        // Prove the authoritative schema rejects both invalid element order and contradictory assertions.
        var title = collections[0].Element(ns + "TitleDetail")!;
        title.Remove();
        collections[0].Add(title);
        RequireRejected(document, schemas);
        title.Remove();
        collections[0].Elements(ns + "Contributor").First().AddBeforeSelf(title);
        collections[0].Add(new XElement(ns + "NoContributor"));
        RequireRejected(document, schemas);
    }

    private static void RequireRejected(XDocument document, XmlSchemaSet schemas) {
        bool rejected = false;
        document.Validate(schemas, (_, args) => { if (args.Severity == XmlSeverityType.Error) rejected = true; });
        if (!rejected) throw new InvalidDataException("The supplied schema accepted invalid collection authorship.");
    }
}
