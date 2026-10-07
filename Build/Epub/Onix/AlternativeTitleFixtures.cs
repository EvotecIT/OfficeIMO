using OfficeIMO.Workflows;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class AlternativeTitleFixtures {
    internal static IReadOnlyList<BookOnixAlternativeTitle> Create() => [
        new() { Type = BookOnixAlternativeTitleType.OriginalLanguage, Title = "L’histoire de l’édition", LanguageCode = "fre", Subtitle = "Une étude", TitleSorting = new() { Prefix = "L’" } },
        new() { Type = BookOnixAlternativeTitleType.Abbreviated, Title = "Publishing", LanguageCode = "eng", TitleSorting = new() },
        new() { Type = BookOnixAlternativeTitleType.OtherLanguage, Title = "出版の歴史", LanguageCode = "jpn" },
        new() { Type = BookOnixAlternativeTitleType.Former, Title = "The printed page", TitleSorting = new() { Prefix = "The " } },
        new() { Type = BookOnixAlternativeTitleType.Distributor, Title = "PUBLISHING HISTORY EPUB" },
        new() { Type = BookOnixAlternativeTitleType.Cover, Title = "Publishing & its history" },
        new() { Type = BookOnixAlternativeTitleType.BackCover, Title = "A history of publishing" },
        new() { Type = BookOnixAlternativeTitleType.Expanded, Title = "Publishing history: an illustrated introduction" },
        new() { Type = BookOnixAlternativeTitleType.Alternative, Title = "The publishing story" },
        new() { Type = BookOnixAlternativeTitleType.Spine, Title = "History of publishing" },
        new() { Type = BookOnixAlternativeTitleType.TranslatedFrom, Title = "Geschichte des Verlagswesens", LanguageCode = "ger" }
    ];

    internal static void Verify(BookOnixExportResult result, XmlSchemaSet schemas) {
        XNamespace ns = BookProject.OnixNamespace;
        var document = XDocument.Load(new MemoryStream(result.Bytes));
        var titles = document.Descendants(ns + "DescriptiveDetail").Single().Elements(ns + "TitleDetail").ToArray();
        if (!titles.Select(t => t.Element(ns + "TitleType")!.Value).SequenceEqual(new[] { "01", "03", "05", "06", "08", "10", "11", "12", "13", "14", "15", "16" }))
            throw new InvalidDataException("Alternative title classifications or order changed.");
        if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
            throw new InvalidDataException("Message composition changed alternative titles.");
        RejectMutation(xml => xml.Descendants(ns + "TitleDetail").Skip(1).First().Element(ns + "TitleType")!.Value = "99");
        RejectMutation(xml => xml.Descendants(ns + "TitlePrefix").First().SetAttributeValue("language", "zzz"));

        void RejectMutation(Action<XDocument> mutate) {
            var invalid = new XDocument(document); mutate(invalid);
            bool rejected = false;
            invalid.Validate(schemas, (_, args) => { if (args.Severity == XmlSeverityType.Error) rejected = true; });
            if (!rejected) throw new InvalidDataException("The supplied schema accepted an invalid alternative title code.");
        }
    }
}
