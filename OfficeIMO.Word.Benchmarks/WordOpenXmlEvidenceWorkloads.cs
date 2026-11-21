using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// The package an SDK caller writes to match OfficeIMO's output: the supporting parts OfficeIMO writes (styles,
/// settings, theme, fonts, notes and properties) and its section layout, captured once outside measurement.
/// </summary>
internal sealed class WordOpenXmlEvidencePackage {
    private readonly string _sectionXml;

    private WordOpenXmlEvidencePackage(WordEquivalentPackageParts parts, string sectionXml) {
        Parts = parts;
        _sectionXml = sectionXml;
    }

    internal WordEquivalentPackageParts Parts { get; }

    /// <summary>Captures the supporting parts and the final section layout of an OfficeIMO package.</summary>
    internal static WordOpenXmlEvidencePackage Capture(byte[] officeImoPackage) {
        using var stream = new MemoryStream(officeImoPackage, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        SectionProperties section = document.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>()?.CloneNode(true) as SectionProperties
            ?? throw new InvalidDataException("The OfficeIMO reference package has no final section.");
        // Header and footer references point at OfficeIMO's own parts; each workload adds its own.
        section.RemoveAllChildren<HeaderReference>();
        section.RemoveAllChildren<FooterReference>();
        return new WordOpenXmlEvidencePackage(WordEquivalentPackageParts.Capture(officeImoPackage), section.OuterXml);
    }

    internal SectionProperties CreateSection() => new(_sectionXml);
}

/// <summary>Format-native authoring of the OfficeIMO creation workloads, creating the package from scratch.</summary>
internal static class WordOpenXmlEvidenceWorkloads {
    private static readonly Lazy<byte[]> RichTemplate = new(CreateRichTemplate);

    internal static byte[] CreateParagraphs(int itemCount, WordOpenXmlEvidencePackage package) {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            package.Parts.AddTo(document, mainPart);
            var body = new Body();
            for (int index = 0; index < itemCount; index++) {
                body.Append(new Paragraph(
                    new ParagraphProperties(),
                    new Run(new Text(WordBenchmarkCorpus.ParagraphText(index)))));
            }
            body.Append(package.CreateSection());
            mainPart.Document = new Document(body);
        }
        return stream.ToArray();
    }

    internal static byte[] CreateReport(int rowCount, WordOpenXmlEvidencePackage package) {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            package.Parts.AddTo(document, mainPart);
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(WordCreateReportComparisonBenchmarks.CreateParagraph(WordBenchmarkCorpus.ReportHeader));
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(WordCreateReportComparisonBenchmarks.CreateParagraph(WordBenchmarkCorpus.ReportFooter));

            SectionProperties section = package.CreateSection();
            section.PrependChild(new FooterReference {
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(footerPart)
            });
            section.PrependChild(new HeaderReference {
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(headerPart)
            });
            // Word sizes a new table without explicit widths to fill the section's text width, column by column.
            var body = new Body(
                WordCreateReportComparisonBenchmarks.CreateParagraph(
                    WordBenchmarkCorpus.ReportTitle,
                    bold: true,
                    fontSizeHalfPoints: "36"),
                WordCreateReportComparisonBenchmarks.CreateParagraph(WordBenchmarkCorpus.ReportSummary),
                WordCreateReportComparisonBenchmarks.CreateTable(rowCount, WordBenchmarkCorpus.SectionTextWidthTwips(section) / 2),
                section);
            mainPart.Document = new Document(body);
        }
        return stream.ToArray();
    }

    internal static byte[] CreateRichReplaceFixture(int itemCount) {
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(RichTemplate.Value);
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            Body body = document.MainDocumentPart?.Document?.Body
                ?? throw new InvalidDataException("The rich Word benchmark template has no body.");
            SectionProperties? section = body.GetFirstChild<SectionProperties>();
            for (int index = 0; index < itemCount; index++) {
                var paragraph = new Paragraph(
                    new ParagraphProperties(),
                    new Run(new Text(WordBenchmarkCorpus.ReplacementText(index))));
                if (section == null) body.Append(paragraph);
                else body.InsertBefore(paragraph, section);
            }
        }
        return stream.ToArray();
    }

    internal static byte[] ReplaceWithOfficeIMO(byte[] fixture) {
        using var input = new MemoryStream(fixture, writable: false);
        using WordDocument document = WordDocument.Load(input);
        document.ReplaceAll(
            WordBenchmarkCorpus.Placeholder,
            WordBenchmarkCorpus.Replacement,
            new global::OfficeIMO.OfficeTextSearchOptions { MatchCase = true });
        return document.ToBytes();
    }

    internal static byte[] ReplaceWithOpenXmlSdk(byte[] fixture) {
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(fixture);
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            MainDocumentPart mainPart = document.MainDocumentPart
                ?? throw new InvalidDataException("The benchmark fixture has no main document part.");
            Body body = mainPart.Document?.Body
                ?? throw new InvalidDataException("The benchmark fixture has no body.");
            foreach (Text text in body.Descendants<Text>()) {
                text.Text = text.Text.Replace(
                    WordBenchmarkCorpus.Placeholder,
                    WordBenchmarkCorpus.Replacement,
                    StringComparison.Ordinal);
            }
            mainPart.Document!.Save();
        }
        return stream.ToArray();
    }

    private static byte[] CreateRichTemplate() {
        using WordDocument document = WordDocument.Create();
        return document.ToBytes();
    }
}

internal sealed class WordRichReplaceEvidenceWorkload {
    internal WordRichReplaceEvidenceWorkload(int itemCount) {
        ItemCount = itemCount;
        Fixture = WordOpenXmlEvidenceWorkloads.CreateRichReplaceFixture(itemCount);
    }

    internal int ItemCount { get; }
    internal byte[] Fixture { get; }
    internal int InputBytes => Fixture.Length;

    internal byte[] OfficeIMO() => WordOpenXmlEvidenceWorkloads.ReplaceWithOfficeIMO(Fixture);
    internal byte[] OpenXmlSdk() => WordOpenXmlEvidenceWorkloads.ReplaceWithOpenXmlSdk(Fixture);
}
