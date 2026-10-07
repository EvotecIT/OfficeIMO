using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class CollectionHierarchyFixtures {
    internal static IReadOnlyList<BookOnixCollection> Create(string profile) => profile switch {
        "collection-hierarchy" => [new() {
            Type = BookOnixCollectionType.Ascribed, SourceName = "Example Library", Frequency = BookOnixCollectionFrequency.Monthly,
            TitleElements = [
                new() { Level = BookOnixCollectionLevel.Subcollection, Title = "Études historiques", Subtitle = "Archives", PartNumber = "Volume II", LanguageCode = "fre" },
                new() { Level = BookOnixCollectionLevel.Collection, Title = "Collected studies", LanguageCode = "eng" },
                new() { Level = BookOnixCollectionLevel.SubSubcollection, PartNumber = "Part 3" }],
            Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "03178471") { Level = BookOnixCollectionLevel.Collection },
                new(BookOnixCollectionIdentifierType.Issn, "1092003X") { Level = BookOnixCollectionLevel.Subcollection },
                new(BookOnixCollectionIdentifierType.Proprietary, "archives-3", "Library catalog") { Level = BookOnixCollectionLevel.SubSubcollection }],
            Sequences = [new(BookOnixCollectionSequenceType.Publication, "2.3.1")],
            Contributors = [new("Series Editor", BookOnixContributorRole.SeriesEditor)]
        }],
        "collection-frequency" => Enum.GetValues<BookOnixCollectionFrequency>().Select(frequency => new BookOnixCollection {
            Type = BookOnixCollectionType.Publisher, Title = "Frequency " + frequency, Frequency = frequency
        }).ToArray(),
        _ => []
    };

    internal static void Verify(BookOnixExportResult result, bool frequencies, XmlSchemaSet schemas) {
        XNamespace ns = BookProject.OnixNamespace;
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        if (frequencies) {
            var values = document.Descendants(ns + "CollectionFrequency").Select(e => e.Value).ToArray();
            if (!values.SequenceEqual(new[] { "u", "i", "r", "e", "a", "b", "t", "q", "s", "m", "f", "w", "d", "x" }))
                throw new InvalidDataException("Collection frequency mappings changed.");
            document.Descendants(ns + "CollectionFrequency").First().Value = "z";
        } else {
            var titles = document.Descendants(ns + "Collection").Single().Descendants(ns + "TitleElement").ToArray();
            if (!titles.Select(e => e.Element(ns + "TitleElementLevel")!.Value).SequenceEqual(new[] { "03", "02", "06" }) ||
                !titles.Select(e => e.Element(ns + "SequenceNumber")!.Value).SequenceEqual(new[] { "1", "2", "3" }) ||
                !document.Descendants(ns + "CollectionElementLevel").Select(e => e.Value).SequenceEqual(new[] { "02", "03", "06" }))
                throw new InvalidDataException("Collection hierarchy or identifier scopes changed.");
            titles[1].Element(ns + "SequenceNumber")!.Value = "1";
        }
        bool rejected = false;
        document.Validate(schemas, (_, args) => { if (args.Severity == XmlSeverityType.Error) rejected = true; });
        if (!rejected) throw new InvalidDataException("The supplied schema accepted invalid collection hierarchy or frequency data.");
    }
}
