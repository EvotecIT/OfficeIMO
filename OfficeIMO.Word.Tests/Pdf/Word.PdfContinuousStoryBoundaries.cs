using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SaveAsPdf_ContinuousBodyIgnoresUnusedStoryEdgeDistance(bool footerDistance, bool oppositeStory) {
        using WordDocument source = WordDocument.Create();
        if (oppositeStory) {
            WordHeaderFooter story = footerDistance ? source.HeaderDefaultOrCreate : source.FooterDefaultOrCreate;
            SetInheritedRunningControlStyle(story.AddParagraph("MARK"));
        }
        SetInheritedRunningControlStyle(source.AddParagraph("S1"));
        WordSection later = source.AddSection(WordSectionBreakType.Continuous);
        if (footerDistance) later.Margins.FooterDistance += 720;
        else later.Margins.HeaderDistance += 720;
        SetInheritedRunningControlStyle(source.AddParagraph("S2"));
        // The public story convenience API can create an empty opposite part.
        // This control specifically has no authored story on the changed edge.
        foreach (WordSection section in source.Sections)
            foreach (var reference in section._sectionProperties.ChildElements.Where(element =>
                footerDistance ? element is W.FooterReference : element is W.HeaderReference).ToArray()) reference.Remove();
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("S1", pdf.GetPage(1).Text);
        Assert.Contains("S2", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ContinuousInheritedStoryPreservesChangedEdgeDistance(bool footer) {
        using WordDocument source = CreateInheritedRunningStoryControl(footer, "default", "continuous");
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = 360;
        source.Sections[1].Margins.HeaderDistance = source.Sections[1].Margins.FooterDistance = 1080;
        source.Paragraphs.Single(p => p.Text == "S2A").PageBreakBefore = true;
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(2, pdf.NumberOfPages);
        double first = pdf.GetPage(1).GetWords().Single(word => word.Text == "MARK").Letters[0].StartBaseLine.Y;
        double second = pdf.GetPage(2).GetWords().Single(word => word.Text == "MARK").Letters[0].StartBaseLine.Y;
        Assert.Equal(footer ? 36D : -36D, second - first, 3);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SaveAsPdf_ContinuousInheritedFirstStoryUsesLogicalSectionPage(bool footer, bool sharedBoundaryPage) {
        using WordDocument source = CreateInheritedRunningStoryControl(footer, "first", "continuous");
        WordParagraph later = source.Paragraphs.Single(p => p.Text == "S2A");
        later.PageBreakBefore = !sharedBoundaryPage;
        later.AddBreak(WordBreakType.Page);
        SetInheritedRunningControlStyle(source.AddParagraph("S2B"));
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(sharedBoundaryPage ? 2 : 3, pdf.NumberOfPages);
        for (int page = 1; page <= pdf.NumberOfPages; page++)
            Assert.Equal(page == 1 || (!sharedBoundaryPage && page == 2), pdf.GetPage(page).Text.Contains("MARK"));
        Assert.Contains("S2A", pdf.GetPage(sharedBoundaryPage ? 1 : 2).Text);
        Assert.Contains("S2B", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void SaveAsPdf_InheritedRequiredStoryDoesNotBypassMixedTextBoxVariant(bool footer, bool even, bool table) {
        using WordDocument source = WordDocument.Create();
        WordHeaderFooter required = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        required.AddParagraph("COUNT ").AddField(WordFieldType.SectionPages);
        if (table) required.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "TABLE";
        SetInheritedRunningControlStyle(source.AddParagraph("S1"));
        WordSection later = source.AddSection();
        later.AddHeadersAndFooters();
        later.DifferentFirstPage = !even;
        later.DifferentOddAndEvenPages = even;
        WordHeaderFooter mixed = footer ? even ? later.Footer!.Even! : later.Footer!.First!
            : even ? later.Header!.Even! : later.Header!.First!;
        foreach (var element in mixed.ChildElements.ToArray()) element.Remove();
        WordParagraph paragraph = mixed.AddParagraph("PREFIX");
        paragraph.AddText(string.Empty).AddTextBox("BOX", WordImageTextWrapping.Square);
        paragraph.AddText("SUFFIX");
        foreach (var reference in later._sectionProperties.ChildElements
            .Where(element => element is W.HeaderReference header && header.Type?.Value == W.HeaderFooterValues.Default ||
                element is W.FooterReference foot && foot.Type?.Value == W.HeaderFooterValues.Default).ToArray()) reference.Remove();
        SetInheritedRunningControlStyle(source.AddParagraph("S2"));
        source.Paragraphs.Last().AddBreak(WordBreakType.Page);
        SetInheritedRunningControlStyle(source.AddParagraph("S3"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        using WordDocument control = WordDocument.Load(new MemoryStream(source.ToBytes()));
        WordHeaderFooter controlDefault = footer ? control.FooterDefaultOrCreate : control.HeaderDefaultOrCreate;
        foreach (var element in controlDefault.ChildElements.ToArray()) element.Remove();
        controlDefault.AddParagraph("COUNT");
        using var controlPdf = PdfPigDocument.Open(control.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(3, pdf.NumberOfPages);
        string visible = string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value));
        int prefix = visible.IndexOf("PREFIX", StringComparison.Ordinal);
        int box = visible.IndexOf("BOX", StringComparison.Ordinal);
        int suffix = visible.IndexOf("SUFFIX", StringComparison.Ordinal);
        Assert.True(prefix >= 0 && box > prefix && suffix > box, visible);
        foreach (string marker in new[] { "PREFIX", "BOX", "SUFFIX" })
            Assert.Equal(1, visible.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        Assert.Equal(controlPdf.GetPage(2).Letters.Select(letter => (letter.Value, letter.StartBaseLine)),
            pdf.GetPage(2).Letters.Select(letter => (letter.Value, letter.StartBaseLine)));
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        if (table) Assert.Contains("TABLE", pdf.GetPage(3).Text);
    }
}
