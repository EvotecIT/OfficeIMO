using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_MultipleSections() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfSections.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfSections.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.Header!.Default.AddParagraph("Header1");
            document.Footer!.Default.AddParagraph("Footer1");
            document.AddParagraph("Section1 Paragraph");

            WordSection section2 = document.AddSection();
            section2.AddHeadersAndFooters();
            section2.Header!.Default.AddParagraph("Header2");
            section2.Footer!.Default.AddParagraph("Footer2");
            document.AddParagraph("Section2 Paragraph");

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_HeaderFooterVariants() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfHeaderVariants.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfHeaderVariants.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.Header!.Default.AddParagraph("DefaultHeader");
            document.Footer!.Default.AddParagraph("DefaultFooter");

            for (int i = 0; i < 100; i++) {
                document.AddParagraph($"Paragraph {i}");
            }

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }
}
