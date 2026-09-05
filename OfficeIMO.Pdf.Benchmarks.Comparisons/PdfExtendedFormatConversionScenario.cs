using System.Text;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Pdf;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfExtendedFormatConversionKind {
    AsciiDoc,
    Latex,
    Mhtml,
    OneNote,
    Odt,
    Ods,
    Odp,
    Visio
}

internal sealed record PdfExtendedFormatConversionScenario(
    PdfExtendedFormatConversionKind Kind,
    byte[] SourceBytes,
    IReadOnlyList<string> RequiredText,
    Func<byte[]> ConvertToPdf) {
    private const string Heading = "EXTENDED FORMAT PDF BENCHMARK";

    internal static PdfExtendedFormatConversionScenario Create(PdfExtendedFormatConversionKind kind) {
        byte[] source = kind switch {
            PdfExtendedFormatConversionKind.AsciiDoc => CreateAsciiDoc(),
            PdfExtendedFormatConversionKind.Latex => CreateLatex(),
            PdfExtendedFormatConversionKind.Mhtml => CreateMhtml(),
            PdfExtendedFormatConversionKind.OneNote => CreateOneNote(),
            PdfExtendedFormatConversionKind.Odt => CreateOdt(),
            PdfExtendedFormatConversionKind.Ods => CreateOds(),
            PdfExtendedFormatConversionKind.Odp => CreateOdp(),
            PdfExtendedFormatConversionKind.Visio => CreateVisio(),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };

        return new PdfExtendedFormatConversionScenario(
            kind,
            source,
            PdfFormatConversionRecordManifest.CreateRequiredText(Heading),
            () => Convert(kind, source));
    }

    private static byte[] Convert(PdfExtendedFormatConversionKind kind, byte[] source) {
        using var stream = new MemoryStream(source, writable: false);
        return kind switch {
            PdfExtendedFormatConversionKind.AsciiDoc =>
                AsciiDocDocument.ParseResult(Encoding.UTF8.GetString(source)).Document.ToPdfBytes(),
            PdfExtendedFormatConversionKind.Latex =>
                LatexDocument.ParseResult(Encoding.UTF8.GetString(source)).Document.ToPdfBytes(),
            PdfExtendedFormatConversionKind.Mhtml => MhtmlDocument.Load(stream).ToPdfBytes(),
            PdfExtendedFormatConversionKind.OneNote => OneNoteSectionReader.Read(stream).ToPdfBytes(),
            PdfExtendedFormatConversionKind.Odt => OdtDocument.Load(stream).ToPdfBytes(),
            PdfExtendedFormatConversionKind.Ods => OdsDocument.Load(stream).ToPdfBytes(),
            PdfExtendedFormatConversionKind.Odp => OdpPresentation.Load(stream).ToPdfBytes(),
            PdfExtendedFormatConversionKind.Visio => VisioDocument.Load(stream).ToPdfBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
    }

    private static byte[] CreateAsciiDoc() {
        var source = new StringBuilder("= ").Append(Heading).Append("\n\n");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            source.Append("* ").Append(PdfFormatConversionRecordManifest.RecordLine(index)).Append('\n');
        }
        return Encoding.UTF8.GetBytes(source.ToString());
    }

    private static byte[] CreateLatex() {
        var source = new StringBuilder("\\documentclass{article}\n\\begin{document}\n\\section{")
            .Append(Heading)
            .Append("}\n\\begin{itemize}\n");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            source.Append("\\item ").Append(PdfFormatConversionRecordManifest.RecordLine(index)).Append('\n');
        }
        return Encoding.UTF8.GetBytes(source.Append("\\end{itemize}\n\\end{document}\n").ToString());
    }

    private static byte[] CreateMhtml() {
        var html = new StringBuilder("<!doctype html><html><head><meta charset='utf-8'></head><body><h1>")
            .Append(Heading)
            .Append("</h1>");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            html.Append("<p>").Append(PdfFormatConversionRecordManifest.RecordLine(index)).Append("</p>");
        }
        var document = new MhtmlDocument(
            html.Append("</body></html>").ToString(),
            contentLocation: "https://benchmark.officeimo.test/report.html",
            subject: Heading);
        return document.ToBytes();
    }

    private static byte[] CreateOneNote() {
        var section = new OneNoteSection { Name = Heading };
        for (int pageIndex = 0; pageIndex < 12; pageIndex++) {
            var page = new OneNotePage { Title = pageIndex == 0 ? Heading : $"Records {pageIndex + 1}" };
            int first = (pageIndex * 10) + 1;
            for (int index = first; index < first + 10; index++) {
                var paragraph = new OneNoteParagraph();
                paragraph.Runs.Add(new OneNoteTextRun { Text = PdfFormatConversionRecordManifest.RecordLine(index) });
                page.DirectContent.Add(paragraph);
            }
            section.Pages.Add(page);
        }
        return section.ToByteArray();
    }

    private static byte[] CreateOdt() {
        OdtDocument document = OdtDocument.Create();
        document.AddHeading(Heading, 1);
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            document.AddParagraph(PdfFormatConversionRecordManifest.RecordLine(index));
        }
        return document.ToBytes();
    }

    private static byte[] CreateOds() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Report");
        sheet.Cell(0, 0).SetString(Heading);
        sheet.Cell(1, 0).SetString("Record");
        sheet.Cell(1, 1).SetString("Customer and description");
        sheet.Cell(1, 2).SetString("Amount");
        sheet.Cell(1, 3).SetString("Status");
        for (int index = 1; index <= PdfFormatConversionRecordManifest.RecordCount; index++) {
            int row = index + 1;
            sheet.Cell(row, 0).SetString(PdfFormatConversionRecordManifest.RecordMarker(index));
            sheet.Cell(row, 1).SetString(PdfFormatConversionRecordManifest.CustomerMarker(index) + " " + PdfFormatConversionRecordManifest.Description);
            sheet.Cell(row, 2).SetString(PdfFormatConversionRecordManifest.AmountMarker(index));
            sheet.Cell(row, 3).SetString(PdfFormatConversionRecordManifest.StatusMarker(index));
        }
        return document.ToBytes();
    }

    private static byte[] CreateOdp() {
        OdpPresentation presentation = OdpPresentation.Create();
        for (int slideIndex = 0; slideIndex < 12; slideIndex++) {
            OdpSlide slide = presentation.AddSlide($"Records {slideIndex + 1}");
            int first = (slideIndex * 10) + 1;
            string content = string.Join("\n", Enumerable.Range(first, 10).Select(PdfFormatConversionRecordManifest.RecordLine));
            slide.AddTextBox(
                OdfRect.FromCentimeters(1.5, 1, 30, 16),
                slideIndex == 0 ? Heading + "\n" + content : content,
                $"records-{slideIndex + 1}");
        }
        return presentation.ToBytes();
    }

    private static byte[] CreateVisio() {
        VisioDocument document = VisioDocument.Create();
        for (int pageIndex = 0; pageIndex < 12; pageIndex++) {
            VisioPage page = document.AddPage($"Records {pageIndex + 1}", 8.5, 11);
            if (pageIndex == 0) {
                page.Shapes.Add(new VisioShape("heading", 4.25, 10.4, 7.5, 0.4, Heading));
            }
            int first = (pageIndex * 10) + 1;
            for (int offset = 0; offset < 10; offset++) {
                int index = first + offset;
                page.Shapes.Add(new VisioShape(
                    $"record-{index}",
                    4.25,
                    9.6 - (offset * 0.85),
                    7.5,
                    0.6,
                    PdfFormatConversionRecordManifest.RecordLine(index)));
            }
        }
        return document.ToBytes();
    }

}
