using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfImportTests {
    [Fact]
    public void PdfBytes_ToRtfDocument_Imports_Metadata_Headings_Lists_Paragraphs_And_PageBreaks() {
        byte[] pdf = CreateSemanticPdf();

        RtfDocument document = pdf.ToRtfDocumentFromPdf(CreateReadOptions());

        Assert.Equal("PDF Import Title", document.Info.Title);
        Assert.Equal("OfficeIMO", document.Info.Author);
        Assert.Equal("PDF to RTF", document.Info.Subject);
        Assert.Equal("pdf,rtf", document.Info.Keywords);
        Assert.Equal(7200, document.PageSetup.PaperWidthTwips);
        Assert.Equal(7200, document.PageSetup.PaperHeightTwips);
        Assert.Contains(document.Styles, style => style.Name == "Heading 1" && style.Bold == true);

        RtfParagraph heading = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        Assert.Equal(0, heading.OutlineLevel);
        Assert.True(heading.Runs[0].Bold);

        Assert.Contains(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("semantic paragraph", StringComparison.Ordinal));

        RtfParagraph bullet = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "First bullet");
        Assert.Equal(RtfListKind.Bullet, bullet.ListKind);
        Assert.Equal(0, bullet.ListLevel);
        Assert.NotNull(bullet.ListText);

        RtfParagraph numbered = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "First numbered");
        Assert.Equal(RtfListKind.Decimal, numbered.ListKind);
        Assert.NotNull(numbered.ListText);

        RtfParagraph secondPage = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");
        Assert.True(secondPage.PageBreakBefore);
    }

    [Fact]
    public void Pdf_ToRtf_Serializes_And_RoundTrips_Imported_Text() {
        byte[] pdf = CreateSemanticPdf();

        string rtf = pdf.ToRtfFromPdf(CreateReadOptions());
        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;

        Assert.Contains("Clinical Summary", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ListKind == RtfListKind.Bullet && paragraph.ToPlainText() == "First bullet");
    }

    [Fact]
    public void Pdf_Stream_File_And_Save_Apis_Import_Rtf() {
        byte[] pdf = CreateSemanticPdf();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-rtf-import-" + Guid.NewGuid().ToString("N"));
        string pdfPath = Path.Combine(directory, "source.pdf");
        string rtfPath = Path.Combine(directory, "output.rtf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(pdfPath, pdf);

            using MemoryStream pdfStream = new MemoryStream(pdf);
            RtfDocument fromStream = pdfStream.ToRtfDocumentFromPdf(CreateReadOptions());
            Assert.Contains(fromStream.Paragraphs, paragraph => paragraph.ToPlainText() == "First bullet");

            RtfDocument fromFile = pdfPath.ToRtfDocumentFromPdfFile(CreateReadOptions());
            Assert.Contains(fromFile.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");

            RtfPdfConverterExtensions.SavePdfFileAsRtf(pdfPath, rtfPath, CreateReadOptions());
            RtfDocument saved = RtfDocument.Load(rtfPath, encoding: Encoding.UTF8).Document;
            Assert.Contains(saved.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfRtfReadOptions_Clone_Does_Not_Share_Layout_Options() {
        var options = new PdfRtfReadOptions {
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true,
                MarginLeft = 12
            },
            PreservePageBreaks = false
        };

        PdfRtfReadOptions clone = options.Clone();
        clone.LayoutOptions!.ForceSingleColumn = false;
        clone.LayoutOptions.MarginLeft = 24;

        Assert.True(options.LayoutOptions!.ForceSingleColumn);
        Assert.Equal(12, options.LayoutOptions.MarginLeft);
        Assert.False(clone.PreservePageBreaks);
    }

    private static PdfRtfReadOptions CreateReadOptions() => new PdfRtfReadOptions {
        LayoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        }
    };

    private static byte[] CreateSemanticPdf() =>
        PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 360,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "PDF Import Title", author: "OfficeIMO", subject: "PDF to RTF", keywords: "pdf,rtf")
            .H1("Clinical Summary")
            .Paragraph(p => p.Text("This semantic paragraph should become one imported RTF paragraph."))
            .Bullets(new[] { "First bullet", "Second bullet" })
            .Numbered(new[] { "First numbered", "Second numbered" }, startNumber: 3)
            .PageBreak()
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();
}
