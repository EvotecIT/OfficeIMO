using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfFormatConversionKind {
    Docx,
    Xlsx,
    Pptx,
    Html,
    Markdown,
    Rtf
}

internal sealed record PdfFormatConversionScenario(
    PdfFormatConversionKind Kind,
    byte[] SourceBytes,
    IReadOnlyList<string> RequiredText,
    Func<byte[]> ConvertToPdf) {
    private const string Heading = "FORMAT PDF BENCHMARK";

    internal static PdfFormatConversionScenario Create(
        PdfFormatConversionKind kind,
        PdfCore.PdfTextFallbackFeatures? textFallbacksOverride = null) {
        byte[] source = kind switch {
            PdfFormatConversionKind.Docx => CreateDocx(),
            PdfFormatConversionKind.Xlsx => CreateXlsx(),
            PdfFormatConversionKind.Pptx => CreatePptx(),
            PdfFormatConversionKind.Html => CreateHtml(),
            PdfFormatConversionKind.Markdown => CreateMarkdown(),
            PdfFormatConversionKind.Rtf => CreateRtf(),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };

        return new PdfFormatConversionScenario(
            kind,
            source,
            PdfFormatConversionRecordManifest.CreateRequiredText(Heading),
            () => Convert(kind, source, textFallbacksOverride));
    }

    private static byte[] Convert(
        PdfFormatConversionKind kind,
        byte[] source,
        PdfCore.PdfTextFallbackFeatures? textFallbacksOverride) {
        using var stream = new MemoryStream(source, writable: false);
        return kind switch {
            PdfFormatConversionKind.Docx => ConvertDocx(stream, textFallbacksOverride),
            PdfFormatConversionKind.Xlsx => ConvertXlsx(stream, textFallbacksOverride),
            PdfFormatConversionKind.Pptx => ConvertPptx(stream, textFallbacksOverride),
            PdfFormatConversionKind.Html => ConvertHtml(source, textFallbacksOverride),
            PdfFormatConversionKind.Markdown => ConvertMarkdown(source, textFallbacksOverride),
            PdfFormatConversionKind.Rtf => RtfDocument.LoadResult(stream).Document.ToPdfBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
    }

    private static byte[] ConvertDocx(Stream stream, PdfCore.PdfTextFallbackFeatures? textFallbacksOverride) {
        using WordDocument document = WordDocument.Load(stream);
        return textFallbacksOverride.HasValue
            ? document.ToPdfBytes(new WordToPdfOptions { TextFallbacks = textFallbacksOverride.Value })
            : document.ToPdfBytes();
    }

    private static byte[] ConvertXlsx(Stream stream, PdfCore.PdfTextFallbackFeatures? textFallbacksOverride) {
        using ExcelDocument document = ExcelDocument.Load(stream);
        return textFallbacksOverride.HasValue
            ? document.ToPdfBytes(new ExcelToPdfOptions { TextFallbacks = textFallbacksOverride.Value })
            : document.ToPdfBytes();
    }

    private static byte[] ConvertPptx(Stream stream, PdfCore.PdfTextFallbackFeatures? textFallbacksOverride) {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(stream);
        return textFallbacksOverride.HasValue
            ? presentation.ToPdfBytes(new PowerPointToPdfOptions { TextFallbacks = textFallbacksOverride.Value })
            : presentation.ToPdfBytes();
    }

    private static byte[] ConvertHtml(byte[] source, PdfCore.PdfTextFallbackFeatures? textFallbacksOverride) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(Encoding.UTF8.GetString(source));
        return textFallbacksOverride.HasValue
            ? document.ToPdfBytes(new HtmlToPdfOptions { TextFallbacks = textFallbacksOverride.Value })
            : document.ToPdfBytes();
    }

    private static byte[] ConvertMarkdown(byte[] source, PdfCore.PdfTextFallbackFeatures? textFallbacksOverride) {
        OfficeIMO.Markdown.MarkdownDoc document = OfficeIMO.Markdown.MarkdownReader.Parse(Encoding.UTF8.GetString(source));
        return textFallbacksOverride.HasValue
            ? document.ToPdfBytes(new MarkdownToPdfOptions { TextFallbacks = textFallbacksOverride.Value })
            : document.ToPdfBytes();
    }

    private static byte[] CreateDocx() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(Heading).SetStyle(WordParagraphStyles.Heading1);
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            document.AddParagraph(PdfFormatConversionRecordManifest.RecordLine(index));
        }
        return document.ToBytes();
    }

    private static byte[] CreateXlsx() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, Heading);
        sheet.Cell(2, 1, "Record");
        sheet.Cell(2, 2, "Description");
        sheet.Cell(2, 3, "Amount");
        sheet.Cell(2, 4, "Status");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            int row = index + 2;
            sheet.Cell(row, 1, PdfFormatConversionRecordManifest.RecordMarker(index));
            sheet.Cell(row, 2, PdfFormatConversionRecordManifest.CustomerMarker(index) + " " + PdfFormatConversionRecordManifest.Description);
            sheet.Cell(row, 3, PdfFormatConversionRecordManifest.AmountMarker(index));
            sheet.Cell(row, 4, PdfFormatConversionRecordManifest.StatusMarker(index));
        }
        sheet.SetColumnWidth(1, 18);
        sheet.SetColumnWidth(2, 42);
        sheet.SetColumnWidth(3, 14);
        sheet.SetColumnWidth(4, 14);
        return document.ToBytes();
    }

    private static byte[] CreatePptx() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        for (int slideIndex = 0; slideIndex < 12; slideIndex++) {
            PowerPointSlide slide = presentation.AddSlide();
            int first = (slideIndex * 10) + 1;
            string content = string.Join("\n", Enumerable.Range(first, 10).Select(PdfFormatConversionRecordManifest.RecordLine));
            slide.AddTextBoxPoints(
                slideIndex == 0 ? Heading + "\n" + content : content,
                leftPoints: 36,
                topPoints: 28,
                widthPoints: 640,
                heightPoints: 460);
        }
        return presentation.ToBytes();
    }

    private static byte[] CreateHtml() {
        var html = new StringBuilder("<!doctype html><html><head><meta charset='utf-8'><style>body{font:10pt sans-serif}p{margin:0 0 4pt}</style></head><body>")
            .Append("<h1>").Append(Heading).Append("</h1>");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            html.Append("<p>").Append(PdfFormatConversionRecordManifest.RecordLine(index)).Append("</p>");
        }
        return Encoding.UTF8.GetBytes(html.Append("</body></html>").ToString());
    }

    private static byte[] CreateMarkdown() {
        var markdown = new StringBuilder("# ").Append(Heading).Append("\n\n");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            markdown.Append("- ").Append(PdfFormatConversionRecordManifest.RecordLine(index)).Append('\n');
        }
        return Encoding.UTF8.GetBytes(markdown.ToString());
    }

    private static byte[] CreateRtf() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph(Heading).SetStyle(1);
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            document.AddParagraph(PdfFormatConversionRecordManifest.RecordLine(index));
        }
        return document.ToBytes();
    }

}
