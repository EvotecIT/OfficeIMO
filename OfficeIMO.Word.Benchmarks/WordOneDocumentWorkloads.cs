#if NET10_0_OR_GREATER
using RecursiveAI.OneDocument;
using RecursiveAI.OneDocument.Documents;
using RecursiveAI.OneDocument.Operations;
using OneDocumentDocument = RecursiveAI.OneDocument.Documents.Document;
using OneDocumentParagraph = RecursiveAI.OneDocument.Documents.Paragraph;
using OneDocumentTable = RecursiveAI.OneDocument.Documents.Table;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// OneDocument implementations of the shared Word comparison workloads. The package targets
/// net10.0 only, so these lanes exist only in the net10.0 build of the benchmark runner.
/// </summary>
internal static class WordOneDocumentWorkloads {
    // The other lanes replace case-sensitively; OneDocument searches case-insensitively unless told otherwise.
    private static readonly DocumentReplaceOptions MatchCase = new() { Search = new SearchOptions { MatchCase = true } };

    internal static WordReadObservation ReadParagraphs(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using OneDocumentDocument document = OneDocumentDocument.Open(stream);
        var observation = WordReadObservation.Empty;
        foreach (OneDocumentParagraph paragraph in document.Paragraphs) {
            observation = observation.Add(paragraph.Text);
        }
        return observation;
    }

    internal static WordReadObservation ExtractText(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using OneDocumentDocument document = OneDocumentDocument.Open(stream);
        string text = document.ExtractText();
        var observation = WordReadObservation.Empty;
        foreach (string line in text.Split('\n')) {
            if (line.Length > 0) {
                observation = observation.Add(line);
            }
        }
        return observation;
    }

    internal static byte[] Replace(byte[] fixture) {
        using var input = new MemoryStream(fixture, writable: false);
        using OneDocumentDocument document = OneDocumentDocument.Open(input);
        document.Replace(WordBenchmarkCorpus.Placeholder, WordBenchmarkCorpus.Replacement, MatchCase);
        using var output = new MemoryStream();
        document.SaveAs(output, OfficeFormat.Docx);
        return output.ToArray();
    }

    internal static byte[] CreateParagraphs(int itemCount) {
        using OneDocumentDocument document = OneDocumentDocument.Create();
        // A new OneDocument document starts with one empty paragraph; it takes the first item.
        document.Paragraphs.First().AddRun(WordBenchmarkCorpus.ParagraphText(0));
        for (int index = 1; index < itemCount; index++) {
            document.Body.AddParagraph(WordBenchmarkCorpus.ParagraphText(index));
        }
        return Save(document);
    }

    internal static byte[] CreateImages(int imageCount) {
        using OneDocumentDocument document = OneDocumentDocument.Create();
        // A new OneDocument document starts with one empty paragraph; it takes the first caption.
        document.Paragraphs.First().AddRun(WordWorkloadCorpus.ImageCaption(0));
        for (int index = 0; index < imageCount; index++) {
            if (index > 0) document.Body.AddParagraph(WordWorkloadCorpus.ImageCaption(index));
            document.Body.AddPicture(WordBenchmarkImages.Image(index), PixelsToPoints(WordBenchmarkImages.WidthPixels),
                PixelsToPoints(WordBenchmarkImages.HeightPixels), altText: null);
        }
        return Save(document);
    }

    internal static byte[] FillTemplate(byte[] fixture) {
        using var input = new MemoryStream(fixture, writable: false);
        using OneDocumentDocument document = OneDocumentDocument.Open(input);
        foreach (KeyValuePair<string, string> field in WordWorkloadCorpus.Fields) {
            document.Replace(field.Key, field.Value, MatchCase);
        }
        return Save(document);
    }

    internal static WordReadObservation ReadTables(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using OneDocumentDocument document = OneDocumentDocument.Open(stream);
        var observation = WordReadObservation.Empty;
        foreach (OneDocumentTable table in document.Tables) {
            foreach (TableRow row in table.Rows()) {
                foreach (TableCell cell in row.Cells()) {
                    observation = observation.Add(cell.Text);
                }
            }
        }
        return observation;
    }

    internal static byte[] CreateReport(int rowCount) {
        using OneDocumentDocument document = OneDocumentDocument.Create();
        Section section = document.Sections[0];
        SetStoryText(section.AddOwn(HeaderFooterKind.Header, HeaderFooterType.Default), WordBenchmarkCorpus.ReportHeader);
        SetStoryText(section.AddOwn(HeaderFooterKind.Footer, HeaderFooterType.Default), WordBenchmarkCorpus.ReportFooter);

        // A new OneDocument document starts with one empty paragraph; it becomes the title.
        OneDocumentParagraph title = document.Paragraphs.First();
        Run titleRun = title.AddRun(WordBenchmarkCorpus.ReportTitle);
        titleRun.Bold = true;
        titleRun.FontSize = 18;
        document.Body.AddParagraph(WordBenchmarkCorpus.ReportSummary);

        var cells = new string?[rowCount + 1, 2];
        cells[0, 0] = "Record";
        cells[0, 1] = "Owner";
        for (int index = 0; index < rowCount; index++) {
            cells[index + 1, 0] = WordBenchmarkCorpus.RecordId(index);
            cells[index + 1, 1] = WordBenchmarkCorpus.RecordOwner(index);
        }
        OneDocumentTable table = document.Body.AddTable(cells);
        table[0, 0].Paragraphs().First().Runs().First().Bold = true;
        table[0, 1].Paragraphs().First().Runs().First().Bold = true;

        using var output = new MemoryStream();
        document.SaveAs(output, OfficeFormat.Docx);
        return output.ToArray();
    }

    private static double PixelsToPoints(int pixels) => pixels * 72.0 / 96.0;

    private static byte[] Save(OneDocumentDocument document) {
        using var output = new MemoryStream();
        document.SaveAs(output, OfficeFormat.Docx);
        return output.ToArray();
    }

    private static void SetStoryText(HeaderFooter story, string text) {
        // AddOwn creates the story with one empty paragraph; the typed view writes the run into it.
        foreach (var paragraph in story.View.Paragraphs) {
            paragraph.AddRun().AddText().Text = text;
            return;
        }
        throw new InvalidDataException("OneDocument created a story without a paragraph.");
    }
}
#endif
