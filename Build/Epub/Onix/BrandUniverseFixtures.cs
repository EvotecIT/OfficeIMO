using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class BrandUniverseFixtures {
    internal static IReadOnlyList<BookOnixCollection> Create() => [
        new() { Type = BookOnixCollectionType.Publisher, TitleElements = [
            new() { Level = BookOnixCollectionLevel.MasterBrand, Title = "The Star Garden", LanguageCode = "eng", TitleSorting = new() { Prefix = "The " } }
        ], Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "star-garden", "Publisher catalog") { Level = BookOnixCollectionLevel.MasterBrand }] },
        new() { Type = BookOnixCollectionType.Ascribed, SourceName = "Example Library", TitleElements = [
            new() { Level = BookOnixCollectionLevel.Universe, Title = "Clockwork City", LanguageCode = "eng", TitleSorting = new() }
        ], Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "clockwork-city", "Library catalog") { Level = BookOnixCollectionLevel.Universe }] },
        new() { Type = BookOnixCollectionType.Publisher, TitleElements = [
            new() { Level = BookOnixCollectionLevel.MasterBrand, Title = "Voyager Tales", LanguageCode = "eng" },
            new() { Level = BookOnixCollectionLevel.Universe, Title = "Orbital Commons", LanguageCode = "eng" },
            new() { Level = BookOnixCollectionLevel.Collection, Title = "Early readers" },
            new() { Level = BookOnixCollectionLevel.Subcollection, Title = "Science adventures" },
            new() { Level = BookOnixCollectionLevel.SubSubcollection, PartNumber = "Part 1" }
        ], Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary, "voyager", "Publisher catalog") { Level = BookOnixCollectionLevel.MasterBrand },
            new(BookOnixCollectionIdentifierType.Proprietary, "orbital", "Publisher catalog") { Level = BookOnixCollectionLevel.Universe }] }
    ];

    internal static void Verify(BookOnixExportResult result, XmlSchemaSet schemas) {
        XNamespace ns = BookProject.OnixNamespace;
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        var collections = document.Descendants(ns + "Collection").ToArray();
        if (collections.Length != 3 || collections[0].Descendants(ns + "TitleElementLevel").Single().Value != "05" ||
            collections[1].Descendants(ns + "TitleElementLevel").Single().Value != "07" ||
            !collections[2].Descendants(ns + "TitleElementLevel").Select(e => e.Value).SequenceEqual(new[] { "05", "07", "02", "03", "06" }) ||
            !document.Descendants(ns + "CollectionElementLevel").Select(e => e.Value).SequenceEqual(new[] { "05", "07", "05", "07" }))
            throw new InvalidDataException("Brand, universe or series title/identifier levels changed.");
        if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
            throw new InvalidDataException("ONIX message composition changed brand or universe assertions.");
        document.Descendants(ns + "CollectionElementLevel").First().Value = "99";
        bool rejected = false;
        document.Validate(schemas, (_, args) => { if (args.Severity == XmlSeverityType.Error) rejected = true; });
        if (!rejected) throw new InvalidDataException("The supplied schema accepted an unknown collection identifier level.");
    }
}
