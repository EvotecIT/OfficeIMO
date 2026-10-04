using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("direct")]
    [InlineData("empty")]
    [InlineData("manual")]
    public void SaveAsPdf_PageBreakOffAfterCoverRetainsTheCoverBoundary(string separator) {
        string path = Path.Combine(_directoryWithFiles, "CoverPageBreakOff-" + separator + ".docx");
        using (WordDocument document = WordDocument.Create(path)) {
            document._document.Body!.Append(CreateNativeCoverPageBlock("OverrideCoverMarker"));
            if (separator == "empty") document.AddParagraph().PageBreakBeforeOverride = false;
            if (separator == "manual") document.AddPageBreak();
            document.AddParagraph("AfterCoverMarker").PageBreakBeforeOverride = false;
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        using PdfPigDocument pdf = PdfPigDocument.Open(reopened.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica"
        }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("OverrideCoverMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("AfterCoverMarker", pdf.GetPage(1).Text);
        Assert.Contains("AfterCoverMarker", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_DirectPageBreakOffOverridesStyle(bool list, bool columns) {
        string path = Path.Combine(_directoryWithFiles, "PageBreakOff.docx");
        using (WordDocument document = WordDocument.Create(path)) {
            const string styleId = "BreakBeforeStyle";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new Style(new StyleName { Val = styleId },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new PageBreakBefore())) {
                    Type = StyleValues.Paragraph, StyleId = styleId, CustomStyle = true
                });
            if (columns) document.Sections[0].ColumnCount = 2;
            document.AddParagraph("BeforeOverride");
            WordParagraph target = list
                ? document.AddList(WordListStyle.Bulleted).AddItem("DirectOffTarget")
                : document.AddParagraph("DirectOffTarget");
            target.SetStyleId(styleId);
            target.PageBreakBeforeOverride = false;
            document.AddParagraph("AfterOverride");
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        byte[] bytes = reopened.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica"
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("BeforeOverride", pdf.GetPage(1).Text);
        Assert.Contains("DirectOffTarget", pdf.GetPage(1).Text);
        Assert.Contains("AfterOverride", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ParagraphPagination_SetTrueReplacesLoadedFalse(bool keepNext) {
        string path = Path.Combine(_directoryWithFiles, "LoadedPaginationOff.docx");
        using (WordDocument document = WordDocument.Create(path)) {
            WordParagraph paragraph = document.AddParagraph("Toggle pagination");
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            if (keepNext) paragraph._paragraph.ParagraphProperties.KeepNext = new KeepNext { Val = false };
            else paragraph._paragraph.ParagraphProperties.KeepLines = new KeepLines { Val = false };
            document.Save();
        }

        using (WordDocument document = WordDocument.Load(path)) {
            WordParagraph paragraph = document.Paragraphs[0];
            if (keepNext) paragraph.KeepWithNext = true;
            else paragraph.KeepLinesTogether = true;
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        Assert.True(keepNext ? reopened.Paragraphs[0].KeepWithNext : reopened.Paragraphs[0].KeepLinesTogether);
    }
}
