namespace OfficeIMO.OpenDocument.Benchmarks;

internal sealed record OpenDocumentBenchmarkScale(string Name, int TextParagraphs, int SpreadsheetRows, int SpreadsheetColumns, int PresentationSlides);

internal static class OpenDocumentBenchmarkCorpus {
    internal static readonly IReadOnlyList<OpenDocumentBenchmarkScale> Scales = new[] {
        new OpenDocumentBenchmarkScale("Small", 100, 100, 8, 5),
        new OpenDocumentBenchmarkScale("Normal", 2_000, 1_000, 8, 30),
        new OpenDocumentBenchmarkScale("Large", 10_000, 5_000, 8, 120)
    };

    internal static OpenDocumentBenchmarkScale Get(string name) =>
        Scales.FirstOrDefault(scale => string.Equals(scale.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown scale '{name}'.", nameof(name));

    internal static byte[] CreatePackage(string format, OpenDocumentBenchmarkScale scale) =>
        format.ToUpperInvariant() switch {
            "ODT" => CreateText(scale),
            "ODS" => CreateSpreadsheet(scale),
            "ODP" => CreatePresentation(scale),
            _ => throw new ArgumentException($"Unknown OpenDocument format '{format}'.", nameof(format))
        };

    private static byte[] CreateText(OpenDocumentBenchmarkScale scale) {
        OdtDocument document = OdtDocument.Create();
        for (var index = 0; index < scale.TextParagraphs; index++) {
            document.AddParagraph($"Paragraph {index:D6} contains stable benchmark text and Unicode café 中.");
        }
        return document.ToBytes();
    }

    private static byte[] CreateSpreadsheet(OpenDocumentBenchmarkScale scale) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (var row = 0; row < scale.SpreadsheetRows; row++) {
            sheet.Cell(row, 0).SetNumber(row + 1D);
            for (var column = 1; column < scale.SpreadsheetColumns; column++) {
                sheet.Cell(row, column).SetString($"R{row:D6}C{column:D2}");
            }
        }
        return document.ToBytes();
    }

    private static byte[] CreatePresentation(OpenDocumentBenchmarkScale scale) {
        OdpPresentation presentation = OdpPresentation.Create();
        for (var index = 0; index < scale.PresentationSlides; index++) {
            OdpSlide slide = presentation.AddSlide($"Slide {index + 1}");
            slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 20, 2), $"Title {index + 1}", "Title");
            slide.AddTextBox(OdfRect.FromCentimeters(1, 4, 14, 6), $"Body {index + 1} with stable benchmark content.", "Body");
            OdpRectangle card = slide.AddRectangle(OdfRect.FromCentimeters(18, 4, 8, 4), "Card");
            card.FillColor = OdfColor.Parse(index % 2 == 0 ? "#D1E9FF" : "#DCFAE6");
        }
        return presentation.ToBytes();
    }
}
