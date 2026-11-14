using System.Globalization;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfOfficeDocumentKind {
    Docx,
    Pptx
}

/// <summary>
/// Measures DOCX and PPTX to PDF at two document sizes, one twice the other, so a
/// revision that makes conversion grow faster than the page or slide count shows up
/// as a ratio change. Word pages carry a header and footer, a heading, mixed run
/// formatting, a table, numbered and bulleted lists and periodic images and charts;
/// slides carry text boxes, shapes, an image and alternating charts and tables.
/// Cleanup reopens the PDF and requires at least the requested page count and every page or slide marker.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfOfficeDocumentScalingBenchmarks {
    private byte[] _sourceBytes = null!;
    private byte[]? _result;

    [Params(PdfOfficeDocumentKind.Docx, PdfOfficeDocumentKind.Pptx)]
    public PdfOfficeDocumentKind Format { get; set; }

    [Params(25, 50)]
    public int Pages { get; set; }

    [GlobalSetup]
    public void Setup() {
        BenchmarkAffinityGuard.Validate();
        byte[] image = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Brand", "Logo-evotec.png"));
        _sourceBytes = Format == PdfOfficeDocumentKind.Docx ? CreateDocx(Pages, image) : CreatePptx(Pages, image);
    }

    [Benchmark]
    public byte[] ConvertToPdf() {
        using var stream = new MemoryStream(_sourceBytes, writable: false);
        if (Format == PdfOfficeDocumentKind.Docx) {
            using WordDocument document = WordDocument.Load(stream);
            using PdfDocument pdf = document.ToPdf();
            return _result = pdf.ToBytes();
        }

        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        using PdfDocument slides = presentation.ToPdf();
        return _result = slides.ToBytes();
    }

    [GlobalCleanup]
    public void Validate() {
        if (_result == null) {
            throw new InvalidDataException($"{Format}/{Pages} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        if (observation.PageCount < Pages) {
            throw new InvalidDataException($"{Format}/{Pages} produced only {observation.PageCount} pages.");
        }

        for (int index = 1; index <= Pages; index++) {
            string marker = PdfBenchmarkValidation.Normalize(CreateMarker(index));
            if (!observation.NormalizedText.Contains(marker, StringComparison.Ordinal)) {
                throw new InvalidDataException($"{Format}/{Pages} did not preserve marker {marker}.");
            }
        }

        Console.WriteLine(
            $"OFFICE_SCALING_PDF_EVIDENCE format={Format} pages={Pages} sourceBytes={_sourceBytes.Length} " +
            $"pdfBytes={_result.Length} pdfPages={observation.PageCount} textLength={observation.TextLength}");
    }

    private static byte[] CreateDocx(int pages, byte[] image) {
        using WordDocument document = WordDocument.Create();
        _ = document.PageSetup.Header;
        document.PageSetup.GetHeader(WordHeaderFooterType.Default)!.AddParagraph("Office scaling benchmark header");
        document.PageSetup.GetFooter(WordHeaderFooterType.Default)!.AddParagraph("Office scaling benchmark footer");
        for (int page = 1; page <= pages; page++) {
            document.AddParagraph(CreateMarker(page) + " section heading").SetStyle(WordParagraphStyles.Heading1);
            WordParagraph body = document.AddParagraph("Body text with ");
            body.AddText("bold").SetBold();
            body.AddText(" and ");
            body.AddText("italic").SetItalic();
            body.AddText(" runs that wrap across the line. Zażółć gęślą jaźń. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            WordTable table = document.AddTable(8, 4);
            for (int row = 0; row < 8; row++) {
                for (int column = 0; column < 4; column++) {
                    table.Rows[row].Cells[column].Paragraphs[0].Text = row == 0
                        ? "Column " + column.ToString(CultureInfo.InvariantCulture)
                        : "R" + row.ToString(CultureInfo.InvariantCulture) + "C" + column.ToString(CultureInfo.InvariantCulture) +
                          " value " + (page * row + column).ToString(CultureInfo.InvariantCulture);
                }
            }

            WordList numbered = document.AddListNumbered();
            for (int item = 1; item <= 5; item++) {
                numbered.AddItem("Numbered item " + item.ToString(CultureInfo.InvariantCulture));
            }

            WordList bullets = document.AddListBulleted();
            bullets.AddItem("First bullet");
            bullets.AddItem("Second bullet");

            if (page % 3 == 0) {
                using var imageStream = new MemoryStream(image, writable: false);
                document.AddParagraph().AddImage(imageStream, "logo.png", global::OfficeIMO.OfficeLength.Pixels(240), global::OfficeIMO.OfficeLength.Pixels(60));
            }

            if (page % 10 == 5) {
                WordChart chart = document.AddChart("Quarterly values", false,
                    global::OfficeIMO.OfficeLength.Pixels(400), global::OfficeIMO.OfficeLength.Pixels(220));
                chart.AddCategories(new List<string> { "Q1", "Q2", "Q3", "Q4" });
                chart.AddBar("Values", new List<int> { 1 + page % 7, 5, 9, 3 }, OfficeColor.Blue);
            }

            document.AddParagraph("Closing paragraph. " + string.Concat(Enumerable.Repeat("More text to lay out. ", 8)));
            if (page < pages) {
                document.AddPageBreak();
            }
        }

        return document.ToBytes();
    }

    private static byte[] CreatePptx(int slides, byte[] image) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        for (int index = 1; index <= slides; index++) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox(CreateMarker(index) + " slide title",
                global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(20),
                global::OfficeIMO.OfficeLength.Points(600), global::OfficeIMO.OfficeLength.Points(40));
            slide.AddTextBox(
                string.Join("\n", Enumerable.Range(1, 6).Select(line =>
                    "Point " + line.ToString(CultureInfo.InvariantCulture) + ": longer text that wraps inside the text box")),
                global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(70),
                global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(300));
            slide.AddShape(OfficePresetShapeType.Rectangle,
                global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(70),
                global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(60));
            slide.AddShape(OfficePresetShapeType.Ellipse,
                global::OfficeIMO.OfficeLength.Points(540), global::OfficeIMO.OfficeLength.Points(70),
                global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(60));
            using (var imageStream = new MemoryStream(image, writable: false)) {
                slide.AddImage(imageStream, OfficeImageFormat.Png,
                    global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(150),
                    global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(50));
            }

            if (index % 2 == 0) {
                slide.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new OfficeSeries("Sales", new[] { 1D + index % 5, 4D, 3D, 6D }),
                        new OfficeSeries("Cost", new[] { 2D, 3D, 1D, 4D })
                    }));
            } else {
                PowerPointTable table = slide.AddTable(5, 3,
                    global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(250),
                    global::OfficeIMO.OfficeLength.Points(280), global::OfficeIMO.OfficeLength.Points(150));
                for (int row = 0; row < 5; row++) {
                    for (int column = 0; column < 3; column++) {
                        table.GetCell(row, column).Text = "r" + row.ToString(CultureInfo.InvariantCulture) +
                            "c" + column.ToString(CultureInfo.InvariantCulture);
                    }
                }
            }
        }

        return presentation.ToBytes();
    }

    private static string CreateMarker(int index) =>
        "OFFICE-PAGE-" + index.ToString("0000", CultureInfo.InvariantCulture);
}
