using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class CollectionIdentifierFixtures {
    internal static IReadOnlyList<BookOnixCollection> Create() => [new() {
        Type = BookOnixCollectionType.Publisher, Title = "Collection identifier examples",
        Identifiers = [
            new(BookOnixCollectionIdentifierType.GermanNationalBibliography, "example-dnb-001"),
            new(BookOnixCollectionIdentifierType.GermanBooksInPrint, "example-vlb-001"),
            new(BookOnixCollectionIdentifierType.Electre, "example-electre-001"),
            new(BookOnixCollectionIdentifierType.Doi, "10.1000/collection-example"),
            new(BookOnixCollectionIdentifierType.Urn, "urn:example:collection/one?=edition=1#series"),
            new(BookOnixCollectionIdentifierType.JapaneseMagazine, "01234"),
            new(BookOnixCollectionIdentifierType.BnfControlNumber, "example-bnf-001"),
            new(BookOnixCollectionIdentifierType.Ark, "https://example.org/ark:/12345/collection"),
            new(BookOnixCollectionIdentifierType.IssnL, "2434-561X")
        ]
    }];

    internal static void Verify(BookOnixExportResult result, XmlSchemaSet schemas) {
        XNamespace ns = BookProject.OnixNamespace;
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        if (!document.Descendants(ns + "CollectionIDType").Select(e => e.Value).SequenceEqual(
            new[] { "03", "04", "05", "06", "22", "27", "29", "35", "38" }))
            throw new InvalidDataException("Collection identifier scheme codes changed.");
        if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
            throw new InvalidDataException("ONIX message composition changed collection identifiers.");
        document.Descendants(ns + "CollectionIDType").First().Value = "99";
        bool rejected = false;
        document.Validate(schemas, (_, args) => { if (args.Severity == XmlSeverityType.Error) rejected = true; });
        if (!rejected) throw new InvalidDataException("The supplied schema accepted an unknown collection identifier scheme.");
    }
}
