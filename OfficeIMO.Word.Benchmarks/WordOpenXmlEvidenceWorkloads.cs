using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Format-native authoring over the same rich package shell used by OfficeIMO creation.</summary>
internal static class WordOpenXmlEvidenceWorkloads {
    private static readonly Lazy<byte[]> RichTemplate = new(CreateRichTemplate);

    internal static byte[] CreateParagraphs(int itemCount) {
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(RichTemplate.Value);
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            Body body = document.MainDocumentPart?.Document?.Body
                ?? throw new InvalidDataException("The rich Word benchmark template has no body.");
            SectionProperties? section = body.GetFirstChild<SectionProperties>();
            for (int index = 0; index < itemCount; index++) {
                var paragraph = new Paragraph(
                    new ParagraphProperties(),
                    new Run(new Text(WordBenchmarkCorpus.ParagraphText(index))));
                if (section == null) body.Append(paragraph);
                else body.InsertBefore(paragraph, section);
            }
        }
        return stream.ToArray();
    }

    internal static byte[] CreateReport(int rowCount) {
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(RichTemplate.Value);
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            MainDocumentPart mainPart = document.MainDocumentPart
                ?? throw new InvalidDataException("The rich Word benchmark template has no main document part.");
            Body body = mainPart.Document?.Body
                ?? throw new InvalidDataException("The rich Word benchmark template has no body.");
            SectionProperties section = body.GetFirstChild<SectionProperties>() ?? body.AppendChild(new SectionProperties());

            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(WordCreateReportComparisonBenchmarks.CreateParagraph(WordBenchmarkCorpus.ReportHeader));
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(WordCreateReportComparisonBenchmarks.CreateParagraph(WordBenchmarkCorpus.ReportFooter));
            section.RemoveAllChildren<HeaderReference>();
            section.RemoveAllChildren<FooterReference>();
            section.PrependChild(new FooterReference {
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(footerPart)
            });
            section.PrependChild(new HeaderReference {
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(headerPart)
            });

            body.InsertBefore(
                WordCreateReportComparisonBenchmarks.CreateParagraph(
                    WordBenchmarkCorpus.ReportTitle,
                    bold: true,
                    fontSizeHalfPoints: "36"),
                section);
            body.InsertBefore(
                WordCreateReportComparisonBenchmarks.CreateParagraph(WordBenchmarkCorpus.ReportSummary),
                section);
            body.InsertBefore(WordCreateReportComparisonBenchmarks.CreateTable(rowCount), section);
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
        document.FindAndReplace(
            WordBenchmarkCorpus.Placeholder,
            WordBenchmarkCorpus.Replacement,
            StringComparison.Ordinal);
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
