using DocumentFormat.OpenXml.Wordprocessing;
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
            var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            var defaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            defaultHeader.AddParagraph("Header1");
            defaultFooter.AddParagraph("Footer1");
            document.AddParagraph("Section1 Paragraph");

            WordSection section2 = document.AddSection();
            section2.AddHeadersAndFooters();
            var section2Header = RequireSectionHeader(document, 1, HeaderFooterValues.Default);
            var section2Footer = RequireSectionFooter(document, 1, HeaderFooterValues.Default);
            section2Header.AddParagraph("Header2");
            section2Footer.AddParagraph("Footer2");
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
            var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("DefaultHeader");
            footer.AddParagraph("DefaultFooter");

            for (int i = 0; i < 100; i++) {
                document.AddParagraph($"Paragraph {i}");
            }

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }
}
