using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void StyledHeaderFooterPageCountsKeepDocumentAndRestartedSectionTotalsSeparate() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        AddCounts(RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph());
        AddCounts(RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph());
        document.AddParagraph("First section first page");
        document.AddPageBreak();
        document.AddParagraph("First section second page");
        WordSection second = document.AddSection();
        second.AddPageNumbering(1, WordNumberFormat.Decimal);
        AddCounts(RequireSectionHeader(document, 1, HeaderFooterValues.Default).AddParagraph());
        AddCounts(RequireSectionFooter(document, 1, HeaderFooterValues.Default).AddParagraph());
        second.AddParagraph("Restarted section first page");
        document.AddPageBreak();
        second.AddParagraph("Restarted section second page");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Equal(2, pdf.GetPage(3).Text.Split(new[] { "Count1/2/4" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(2, pdf.GetPage(4).Text.Split(new[] { "Count2/2/4" }, StringSplitOptions.None).Length - 1);
        Assert.Empty(document.ValidateDocument());

        static void AddCounts(WordParagraph paragraph) {
            paragraph.Text = "Count";
            foreach (string instruction in new[] { "PAGE", "SECTIONPAGES", "NUMPAGES" }) {
                if (instruction != "PAGE") paragraph._paragraph.Append(new Run(new Text("/")));
                paragraph._paragraph.Append(new SimpleField(new Run(new RunProperties(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new Spacing { Val = 20 },
                    new CharacterScale { Val = 150 }, new FontSize { Val = "24" }), new Text("999"))) {
                    Instruction = " " + instruction + " "
                });
            }
        }
    }
}
