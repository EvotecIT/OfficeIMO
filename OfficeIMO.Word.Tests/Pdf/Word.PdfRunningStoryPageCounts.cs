using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    // Native Word counts a shared boundary page in both sections, while its
    // running story still uses the section whose body begins that page.
    [Theory]
    [InlineData(false, "shared")]
    [InlineData(true, "shared")]
    [InlineData(false, "new-page")]
    [InlineData(true, "new-page")]
    [InlineData(false, "both-two")]
    [InlineData(true, "both-two")]
    [InlineData(false, "empty-first")]
    [InlineData(true, "empty-first")]
    public void SaveAsPdf_ContinuousRunningStoryPreservesSectionPageCounts(bool footer, string boundary) {
        using WordDocument source = WordDocument.Create();
        WordHeaderFooter story = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        WordParagraph field = story.AddParagraph("COUNT ");
        SetInheritedRunningControlStyle(field);
        field.AddField(WordFieldType.SectionPages);
        if (boundary != "empty-first") {
            WordParagraph first = source.AddParagraph("S1A"); SetInheritedRunningControlStyle(first); first.AddBreak(WordBreakType.Page);
            SetInheritedRunningControlStyle(source.AddParagraph("S1B"));
        }
        source.AddSection(WordSectionBreakType.Continuous);
        WordParagraph second = source.AddParagraph("S2A"); SetInheritedRunningControlStyle(second);
        if (boundary == "new-page") second._paragraph.ParagraphProperties!.AddChild(new W.PageBreakBefore(), true);
        if (boundary is "both-two" or "empty-first") {
            second.AddBreak(WordBreakType.Page);
            SetInheritedRunningControlStyle(source.AddParagraph("S2B"));
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(boundary is "shared" or "empty-first" ? 2 : 3, pdf.NumberOfPages);
        Assert.Contains(boundary == "empty-first" ? "COUNT 1" : "COUNT 2", pdf.GetPage(1).Text);
        Assert.Contains("COUNT 2", pdf.GetPage(2).Text);
        if (pdf.NumberOfPages == 3)
            Assert.Contains(boundary == "new-page" ? "COUNT 1" : "COUNT 2", pdf.GetPage(3).Text);
        Assert.Contains("S2A", pdf.GetPage(boundary == "empty-first" ? 1 : boundary == "new-page" ? 3 : 2).Text);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SaveAsPdf_ContinuousRunningStoryCountsBlankBodyFlow(bool footer, bool explicitBreak) {
        using WordDocument source = WordDocument.Create();
        WordHeaderFooter story = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        WordParagraph field = story.AddParagraph("COUNT ");
        SetInheritedRunningControlStyle(field);
        field.AddField(WordFieldType.SectionPages);
        SetInheritedRunningControlStyle(source.AddParagraph("S1A"));
        source.AddSection(WordSectionBreakType.Continuous);
        if (explicitBreak) source.AddParagraph().AddBreak(WordBreakType.Page);
        else for (int index = 0; index < 60; index++)
            SetInheritedRunningControlStyle(source.AddParagraph());
        SetInheritedRunningControlStyle(source.AddParagraph("S2A"));

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("S1A", pdf.GetPage(1).Text);
        Assert.Contains("COUNT 1", pdf.GetPage(1).Text);
        Assert.Contains("S2A", pdf.GetPage(2).Text);
        Assert.Contains("COUNT 2", pdf.GetPage(2).Text);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }
}
