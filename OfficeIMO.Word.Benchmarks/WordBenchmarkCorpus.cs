using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Benchmarks;

internal static class WordBenchmarkCorpus {
    private static readonly XNamespace OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal const string ReportHeader = "OfficeIMO Word benchmark";
    internal const string ReportTitle = "Quarterly benchmark report";
    internal const string ReportSummary = "Equivalent structured report workload";
    internal const string ReportFooter = "Generated without Microsoft Office";
    internal const string Placeholder = "{{Status}}";
    internal const string Replacement = "Approved";

    internal static string ParagraphText(int index) =>
        "Paragraph " + index.ToString("D6", System.Globalization.CultureInfo.InvariantCulture) +
        " contains deterministic benchmark text.";

    internal static string RecordId(int index) =>
        "Record " + index.ToString("D6", System.Globalization.CultureInfo.InvariantCulture);

    internal static string RecordOwner(int index) =>
        "Owner " + (index % 17).ToString("D2", System.Globalization.CultureInfo.InvariantCulture);

    internal static string ReplacementText(int index) =>
        RecordId(index) + " status: " + Placeholder;

    internal static byte[] CreateParagraphFixture(int itemCount, bool withPlaceholder = false) {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
                   stream,
                   WordprocessingDocumentType.Document,
                   autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(
                    new StyleName { Val = "Normal" }) {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    Default = true
                });
            var body = new Body();
            for (int index = 0; index < itemCount; index++) {
                body.Append(new Paragraph(new Run(new Text(
                    withPlaceholder ? ReplacementText(index) : ParagraphText(index)))));
            }
            mainPart.Document = new Document(body);
        }
        return stream.ToArray();
    }

    internal static WordReadObservation ObserveExpectedParagraphs(int itemCount) {
        var observation = WordReadObservation.Empty;
        for (int index = 0; index < itemCount; index++) {
            observation = observation.Add(ParagraphText(index));
        }
        return observation;
    }

    internal static void ValidateParagraphDocument(
        byte[] payload,
        int itemCount,
        bool requireOpenXmlSdkConformance = true) {
        using ZipArchive package = OpenPackage(payload);
        XElement body = LoadMainBody(package);
        XElement[] paragraphs = body.Elements(WordNamespace + "p").ToArray();
        EnsureEqual("body paragraph count", itemCount, paragraphs.Length);
        for (int index = 0; index < itemCount; index++) {
            EnsureEqual("paragraph " + index, ParagraphText(index), ReadText(paragraphs[index]));
        }
        if (requireOpenXmlSdkConformance) EnsureOpenXmlSdkConformance(payload);
    }

    internal static void ValidateReportDocument(
        byte[] payload,
        int rowCount,
        bool requireOpenXmlSdkConformance = true,
        bool requireOfficeCompatibleDefaults = false) {
        using ZipArchive package = OpenPackage(payload);
        XElement body = LoadMainBody(package);
        XElement[] reportElements = body.Elements()
            .Where(element => element.Name != WordNamespace + "sectPr")
            .ToArray();
        if (reportElements.Length < 3 ||
            reportElements[0].Name != WordNamespace + "p" ||
            reportElements[1].Name != WordNamespace + "p" ||
            reportElements[2].Name != WordNamespace + "tbl") {
            throw new InvalidDataException("The report must contain title, summary, and table elements in that order.");
        }

        XElement title = reportElements[0];
        EnsureEqual("report title", ReportTitle, ReadText(title));
        if (!title.Descendants(WordNamespace + "b").Any()) {
            throw new InvalidDataException("The report title is not bold.");
        }
        string? titleSize = title.Descendants(WordNamespace + "sz")
            .Select(element => (string?)element.Attribute(WordNamespace + "val"))
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        EnsureEqual("report title font size", "36", titleSize ?? string.Empty);
        EnsureEqual("report summary", ReportSummary, ReadText(reportElements[1]));

        XElement table = reportElements[2];
        if (requireOfficeCompatibleDefaults) ValidateOfficeCompatibleTableDefaults(table);
        EnsureEqual("report table count", 1, body.Elements(WordNamespace + "tbl").Count());
        XElement[] rows = table.Elements(WordNamespace + "tr").ToArray();
        EnsureEqual("report table row count", rowCount + 1, rows.Length);
        EnsureCell(rows[0], 0, "Record", requireBold: true);
        EnsureCell(rows[0], 1, "Owner", requireBold: true);
        for (int index = 0; index < rowCount; index++) {
            EnsureCell(rows[index + 1], 0, RecordId(index));
            EnsureCell(rows[index + 1], 1, RecordOwner(index));
        }

        XElement sectionProperties = body.Descendants(WordNamespace + "sectPr").LastOrDefault()
            ?? throw new InvalidDataException("The report has no section properties.");
        ValidateStoryReference(
            package,
            sectionProperties,
            "headerReference",
            "header",
            "hdr",
            "report header",
            ReportHeader);
        ValidateStoryReference(
            package,
            sectionProperties,
            "footerReference",
            "footer",
            "ftr",
            "report footer",
            ReportFooter);
        if (requireOpenXmlSdkConformance) EnsureOpenXmlSdkConformance(payload);
    }

    private static void ValidateOfficeCompatibleTableDefaults(XElement table) {
        XElement properties = table.Element(WordNamespace + "tblPr")
            ?? throw new InvalidDataException("The report table has no properties.");
        EnsureEqual(
            "report table style",
            "TableGrid",
            (string?)properties.Element(WordNamespace + "tblStyle")?.Attribute(WordNamespace + "val") ?? string.Empty);
        XElement width = properties.Element(WordNamespace + "tblW")
            ?? throw new InvalidDataException("The report table has no preferred width.");
        EnsureEqual("report table width type", "auto", (string?)width.Attribute(WordNamespace + "type") ?? string.Empty);
        EnsureEqual("report table width", "0", (string?)width.Attribute(WordNamespace + "w") ?? string.Empty);
        EnsureEqual(
            "report table look",
            "04A0",
            (string?)properties.Element(WordNamespace + "tblLook")?.Attribute(WordNamespace + "val") ?? string.Empty);

        XElement[] gridColumns = table.Element(WordNamespace + "tblGrid")?
            .Elements(WordNamespace + "gridCol")
            .ToArray() ?? [];
        EnsureEqual("report table grid column count", 2, gridColumns.Length);
        foreach (XElement column in gridColumns) {
            EnsureEqual("report table grid column width", "2400", (string?)column.Attribute(WordNamespace + "w") ?? string.Empty);
        }

        foreach (XElement cell in table.Elements(WordNamespace + "tr").SelectMany(row => row.Elements(WordNamespace + "tc"))) {
            XElement cellWidth = cell.Element(WordNamespace + "tcPr")?.Element(WordNamespace + "tcW")
                ?? throw new InvalidDataException("A report table cell has no preferred width.");
            EnsureEqual("report cell width type", "dxa", (string?)cellWidth.Attribute(WordNamespace + "type") ?? string.Empty);
            EnsureEqual("report cell width", "2400", (string?)cellWidth.Attribute(WordNamespace + "w") ?? string.Empty);
            if (cell.Element(WordNamespace + "p")?.Element(WordNamespace + "pPr") == null) {
                throw new InvalidDataException("A report table cell has no editable paragraph properties.");
            }
        }
    }

    internal static void ValidateReplacedDocument(
        byte[] payload,
        int itemCount,
        bool requireOpenXmlSdkConformance = true) {
        using ZipArchive package = OpenPackage(payload);
        XElement body = LoadMainBody(package);
        XElement[] paragraphs = body.Elements(WordNamespace + "p").ToArray();
        EnsureEqual("replaced paragraph count", itemCount, paragraphs.Length);
        for (int index = 0; index < itemCount; index++) {
            string expected = ReplacementText(index).Replace(Placeholder, Replacement, StringComparison.Ordinal);
            EnsureEqual("replaced paragraph " + index, expected, ReadText(paragraphs[index]));
        }
        if (requireOpenXmlSdkConformance) EnsureOpenXmlSdkConformance(payload);
    }

    internal static int CountStyleDefinitions(byte[] payload) {
        using ZipArchive package = OpenPackage(payload);
        ZipArchiveEntry stylesEntry = package.GetEntry("word/styles.xml")
            ?? throw new InvalidDataException("The DOCX package has no style definitions part.");
        return LoadXml(stylesEntry).Root?.Elements(WordNamespace + "style").Count() ?? 0;
    }

    internal static MemoryStream CreateEditableStream(byte[] payload) {
        var stream = new MemoryStream(payload.Length * 2);
        stream.Write(payload, 0, payload.Length);
        stream.Position = 0;
        return stream;
    }

    private static void EnsureCell(XElement row, int index, string expected, bool requireBold = false) {
        XElement[] cells = row.Elements(WordNamespace + "tc").ToArray();
        if (index >= cells.Length) {
            throw new InvalidDataException("The report table row has fewer cells than expected.");
        }
        EnsureEqual("report table cell", expected, ReadText(cells[index]));
        if (requireBold && !cells[index].Descendants(WordNamespace + "b").Any()) {
            throw new InvalidDataException("The report table header cell '" + expected + "' is not bold.");
        }
    }

    private static ZipArchive OpenPackage(byte[] payload) {
        var package = new ZipArchive(new MemoryStream(payload, writable: false), ZipArchiveMode.Read);
        if (package.GetEntry("[Content_Types].xml") is null ||
            package.GetEntry("_rels/.rels") is null ||
            package.GetEntry("word/document.xml") is null) {
            package.Dispose();
            throw new InvalidDataException("The generated payload is not a complete DOCX package.");
        }
        return package;
    }

    private static XElement LoadMainBody(ZipArchive package) {
        XDocument document = LoadXml(package.GetEntry("word/document.xml")!);
        return document.Root?.Element(WordNamespace + "body")
            ?? throw new InvalidDataException("The DOCX payload has no document body.");
    }

    private static void ValidateStoryReference(
        ZipArchive package,
        XElement sectionProperties,
        string referenceName,
        string relationshipKind,
        string expectedRootName,
        string validationName,
        string expectedText) {
        XElement reference = sectionProperties.Elements(WordNamespace + referenceName).LastOrDefault()
            ?? throw new InvalidDataException("The report section has no active " + referenceName + ".");
        string relationshipId = (string?)reference.Attribute(OfficeRelationshipNamespace + "id")
            ?? throw new InvalidDataException("The report " + referenceName + " has no relationship id.");

        ZipArchiveEntry relationshipsEntry = package.GetEntry("word/_rels/document.xml.rels")
            ?? throw new InvalidDataException("The DOCX package has no main-document relationships part.");
        XDocument relationships = LoadXml(relationshipsEntry);
        XElement relationship = relationships.Root?
            .Elements(PackageRelationshipNamespace + "Relationship")
            .SingleOrDefault(element => string.Equals(
                (string?)element.Attribute("Id"),
                relationshipId,
                StringComparison.Ordinal))
            ?? throw new InvalidDataException("The report " + referenceName + " relationship cannot be resolved.");
        string expectedRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + relationshipKind;
        EnsureEqual(
            "report " + referenceName + " relationship type",
            expectedRelationshipType,
            (string?)relationship.Attribute("Type") ?? string.Empty);
        if (string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException("The report " + referenceName + " unexpectedly targets an external resource.");
        }

        string target = (string?)relationship.Attribute("Target")
            ?? throw new InvalidDataException("The report " + referenceName + " relationship has no target.");
        var packageBaseUri = new Uri("http://package/word/document.xml", UriKind.Absolute);
        string entryPath = new Uri(packageBaseUri, target).AbsolutePath.TrimStart('/');
        ZipArchiveEntry storyEntry = package.GetEntry(entryPath)
            ?? throw new InvalidDataException("The report " + referenceName + " target '" + entryPath + "' is missing.");
        XElement storyRoot = LoadXml(storyEntry).Root
            ?? throw new InvalidDataException(entryPath + " has no XML root.");
        if (storyRoot.Name != WordNamespace + expectedRootName) {
            throw new InvalidDataException(
                "The report " + referenceName + " target has root '" + storyRoot.Name +
                "'; expected '" + (WordNamespace + expectedRootName) + "'.");
        }
        EnsureEqual(validationName, expectedText, ReadText(storyRoot));
    }

    private static XDocument LoadXml(ZipArchiveEntry entry) {
        using Stream stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static string ReadText(XElement element) =>
        string.Concat(element.Descendants(WordNamespace + "t").Select(text => text.Value));

    private static void EnsureOpenXmlSdkConformance(byte[] payload) {
        using var stream = new MemoryStream(payload, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        ValidationErrorInfo? error = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(document)
            .FirstOrDefault();
        if (error is not null) {
            throw new InvalidDataException(
                "The generated DOCX failed Open XML validation: " + error.Description);
        }
    }

    private static void EnsureEqual(string name, int expected, int actual) {
        if (actual != expected) {
            throw new InvalidDataException(name + " was " + actual + "; expected " + expected + ".");
        }
    }

    private static void EnsureEqual(string name, string expected, string actual) {
        if (!string.Equals(expected, actual, StringComparison.Ordinal)) {
            throw new InvalidDataException(name + " was '" + actual + "'; expected '" + expected + "'.");
        }
    }
}

public readonly record struct WordReadObservation(int ParagraphCount, int CharacterCount, ulong Checksum) {
    public static WordReadObservation Empty => new(0, 0, 14695981039346656037UL);

    public WordReadObservation Add(string text) {
        ulong checksum = Checksum;
        foreach (char character in text) {
            checksum ^= character;
            checksum *= 1099511628211UL;
        }
        checksum ^= 0xFF;
        checksum *= 1099511628211UL;
        return new WordReadObservation(ParagraphCount + 1, CharacterCount + text.Length, checksum);
    }
}
