using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Inputs and oracles for the image, template-fill, and table-read comparison workloads.</summary>
internal static class WordWorkloadCorpus {
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal static readonly string[] Templates = {
        "Dear {{FirstName}} {{LastName}}, thank you for order {{OrderId}}.",
        "Your order ships on {{ShipDate}} to {{City}}, {{Country}}.",
        "Total due: {{Amount}} {{Currency}} by {{DueDate}}.",
        "Questions? Contact {{Agent}} quoting {{OrderId}}."
    };

    internal static readonly KeyValuePair<string, string>[] Fields = {
        new("{{FirstName}}", "Ada"),
        new("{{LastName}}", "Lovelace"),
        new("{{OrderId}}", "SO-48213"),
        new("{{ShipDate}}", "2026-11-14"),
        new("{{City}}", "Krakow"),
        new("{{Country}}", "Poland"),
        new("{{Amount}}", "1,284.50"),
        new("{{Currency}}", "EUR"),
        new("{{DueDate}}", "2026-12-01"),
        new("{{Agent}}", "Grace Hopper")
    };

    internal static string TemplateText(int index) => Templates[index % Templates.Length];

    internal static string FilledText(int index) {
        string text = TemplateText(index);
        foreach (KeyValuePair<string, string> field in Fields) {
            text = text.Replace(field.Key, field.Value, StringComparison.Ordinal);
        }
        return text;
    }

    internal static string ImageCaption(int index) =>
        "Figure " + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>A minimally styled package with one single-run template paragraph per item.</summary>
    internal static byte[] CreateTemplateFixture(int itemCount) {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
                   stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            var body = new Body();
            for (int index = 0; index < itemCount; index++) {
                body.Append(new Paragraph(new Run(new Text(TemplateText(index)))));
            }
            mainPart.Document = new Document(body);
        }
        return stream.ToArray();
    }

    /// <summary>The Open XML SDK report output, used as a neutral table-read input.</summary>
    internal static byte[] CreateReportFixture(int rowCount) {
        var benchmark = new WordCreateReportComparisonBenchmarks { RowCount = rowCount };
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(benchmark.OpenXmlSdk());
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            // DocX requires a style catalog with a default paragraph style before it can enumerate content.
            document.MainDocumentPart!.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(new StyleName { Val = "Table Grid" }) { Type = StyleValues.Table, StyleId = "TableGrid" });
        }
        byte[] payload = stream.ToArray();
        WordBenchmarkCorpus.ValidateReportDocument(payload, rowCount);
        return payload;
    }

    internal static WordReadObservation ExpectedReportCells(int rowCount) {
        WordReadObservation observation = WordReadObservation.Empty.Add("Record").Add("Owner");
        for (int index = 0; index < rowCount; index++) {
            observation = observation.Add(WordBenchmarkCorpus.RecordId(index)).Add(WordBenchmarkCorpus.RecordOwner(index));
        }
        return observation;
    }

    internal static void ValidateFilledDocument(byte[] payload, int itemCount, bool requireOpenXmlSdkConformance = true) {
        using ZipArchive package = new(new MemoryStream(payload, writable: false), ZipArchiveMode.Read);
        XElement body = LoadBody(package);
        XElement[] paragraphs = body.Elements(W + "p").ToArray();
        Ensure("filled paragraph count", itemCount.ToString(), paragraphs.Length.ToString());
        for (int index = 0; index < itemCount; index++) {
            Ensure("filled paragraph " + index, FilledText(index), Text(paragraphs[index]));
        }
        if (requireOpenXmlSdkConformance) WordBenchmarkCorpus.EnsureConformance(payload);
    }

    /// <summary>Requires one caption paragraph and one inline picture paragraph per image, each picture carrying its exact bytes.</summary>
    internal static void ValidateImageDocument(byte[] payload, int imageCount, bool requireOpenXmlSdkConformance = true) {
        using ZipArchive package = new(new MemoryStream(payload, writable: false), ZipArchiveMode.Read);
        XElement body = LoadBody(package);
        XDocument relationships = XDocument.Load(package.GetEntry("word/_rels/document.xml.rels")!.Open());
        Dictionary<string, string> targets = relationships.Root!.Elements(PackageRelationships + "Relationship")
            .ToDictionary(element => (string)element.Attribute("Id")!, element => (string)element.Attribute("Target")!);
        XElement[] paragraphs = body.Elements(W + "p").Where(paragraph => paragraph.Descendants().Any()).ToArray();
        XElement[] captions = paragraphs.Where(paragraph => !paragraph.Descendants(A + "blip").Any() && Text(paragraph).Length > 0).ToArray();
        XElement[] blips = body.Descendants(A + "blip").ToArray();
        Ensure("caption count", imageCount.ToString(), captions.Length.ToString());
        Ensure("picture count", imageCount.ToString(), blips.Length.ToString());
        for (int index = 0; index < imageCount; index++) {
            Ensure("caption " + index, ImageCaption(index), Text(captions[index]));
            string relationshipId = (string?)blips[index].Attribute(R + "embed")
                ?? throw new InvalidDataException("Picture " + index + " is not embedded.");
            string target = targets[relationshipId];
            string entryName = new Uri(new Uri("http://package/word/document.xml"), target).AbsolutePath.TrimStart('/');
            ZipArchiveEntry entry = package.GetEntry(entryName)
                ?? throw new InvalidDataException("Picture " + index + " target '" + entryName + "' is missing.");
            using Stream stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            if (!copy.ToArray().AsSpan().SequenceEqual(WordBenchmarkImages.Image(index))) {
                throw new InvalidDataException("Picture " + index + " does not carry its source image bytes.");
            }
        }
        if (requireOpenXmlSdkConformance) WordBenchmarkCorpus.EnsureConformance(payload);
    }

    private static XElement LoadBody(ZipArchive package) {
        using Stream stream = package.GetEntry("word/document.xml")!.Open();
        return XDocument.Load(stream).Root!.Element(W + "body")
            ?? throw new InvalidDataException("The DOCX payload has no body.");
    }

    private static string Text(XElement element) =>
        string.Concat(element.Descendants(W + "t").Select(text => text.Value));

    private static void Ensure(string name, string expected, string actual) {
        if (!string.Equals(expected, actual, StringComparison.Ordinal)) {
            throw new InvalidDataException(name + " was '" + actual + "'; expected '" + expected + "'.");
        }
    }
}
