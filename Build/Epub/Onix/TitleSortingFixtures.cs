using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class TitleSortingFixtures {
    internal static IReadOnlyList<BookOnixCollection> Collections() => [
        new() { Type = BookOnixCollectionType.Publisher, Title = "The archive", Subtitle = "Collected studies",
            LanguageCode = "eng", TitleSorting = new() { Prefix = "The " } },
        new() { Type = BookOnixCollectionType.Publisher, TitleElements = [
            new() { Level = BookOnixCollectionLevel.Collection, Title = "L’histoire", LanguageCode = "fre", TitleSorting = new() { Prefix = "L’" } },
            new() { Level = BookOnixCollectionLevel.Subcollection, Title = "Chronicles", PartNumber = "Series II", LanguageCode = "eng", TitleSorting = new() },
            new() { Level = BookOnixCollectionLevel.SubSubcollection, PartNumber = "Part 3" }
        ] }
    ];

    internal static void Verify(BookOnixExportResult result, bool noPrefix, XmlSchemaSet schemas) {
        XNamespace ns = BookProject.OnixNamespace;
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        var productTitle = document.Descendants(ns + "DescriptiveDetail").Single().Elements(ns + "TitleDetail").Single().Element(ns + "TitleElement")!;
        if (document.Descendants(ns + "TitleText").Any() || productTitle.Element(ns + "TitleWithoutPrefix") == null)
            throw new InvalidDataException("Explicit title sorting was not serialized.");
        if (noPrefix) {
            if (productTitle.Element(ns + "NoPrefix") == null || productTitle.Element(ns + "TitlePrefix") != null)
                throw new InvalidDataException("Explicit no-prefix assertion was not serialized.");
        } else if ((string?)productTitle.Element(ns + "TitlePrefix") != "The " ||
            (string?)productTitle.Element(ns + "TitleWithoutPrefix") != "history & future") {
            throw new InvalidDataException("The selected product title was not split exactly.");
        }
        if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
            throw new InvalidDataException("ONIX message composition changed title sorting data.");
        // NoPrefix and TitlePrefix are mutually exclusive even when the title remainder is valid.
        if (noPrefix) productTitle.Element(ns + "NoPrefix")!.AddAfterSelf(new XElement(ns + "TitlePrefix", "The "));
        else productTitle.Element(ns + "TitlePrefix")!.AddBeforeSelf(new XElement(ns + "NoPrefix"));
        bool rejected = false;
        document.Validate(schemas, (_, args) => { if (args.Severity == XmlSeverityType.Error) rejected = true; });
        if (!rejected) throw new InvalidDataException("The supplied schema accepted contradictory title sorting declarations.");
    }
}
