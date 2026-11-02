using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void SaveAsPdf_RequiredSiblingPreservesImplicitVariantFrame(bool footer, int emptyVariant) {
        using WordDocument document = CreateEmptyRunningStoryDocument("absent", footer);
        WordHeaderFooter[] stories = footer
            ? new WordHeaderFooter[] { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate }
            : new WordHeaderFooter[] { document.HeaderFirstOrCreate, document.HeaderEvenOrCreate, document.HeaderDefaultOrCreate };
        foreach (var child in stories[emptyVariant].ChildElements.ToArray()) child.Remove();
        for (int index = 0; index < stories.Length; index++) {
            if (index != emptyVariant) stories[index].AddParagraph("STORY" + index);
            WordParagraph body = document.AddParagraph(index == emptyVariant
                ? string.Join("\n", Enumerable.Repeat("A", 60)) : "A");
            SetInheritedRunningControlStyle(body);
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        WordPdfExportOptions options = EmptyRunningOptions();
        options.PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions();
        if (footer) options.PdfOptions.FooterOffsetY = global::OfficeIMO.OfficeLength.Points(24D);
        else options.PdfOptions.HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(24D);
        using var ordinary = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        for (int index = 0; index < stories.Length; index++)
            if (index != emptyVariant) stories[index].Paragraphs.Last().AddText(" ").AddField(WordFieldType.SectionPages);
        using var required = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(ordinary.GetPage(emptyVariant + 1).Letters.Where(letter => letter.Value == "A").Select(letter => letter.StartBaseLine),
            required.GetPage(emptyVariant + 1).Letters.Where(letter => letter.Value == "A").Select(letter => letter.StartBaseLine));
    }

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
        if (footerDistance) later.Margins.FooterDistance+= global::OfficeIMO.OfficeLength.Points(36D);
        else later.Margins.HeaderDistance+= global::OfficeIMO.OfficeLength.Points(36D);
        SetInheritedRunningControlStyle(source.AddParagraph("S2"));
        // The public story convenience API can create an empty opposite part.
        // This control specifically has no authored story on the changed edge.
        foreach (WordSection section in source.Sections)
            foreach (var reference in section._sectionProperties.ChildElements.Where(element =>
                footerDistance ? element is W.FooterReference : element is W.HeaderReference).ToArray()) reference.Remove();
        using var pdf = PdfPigDocument.Open(source.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("S1", pdf.GetPage(1).Text);
        Assert.Contains("S2", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ContinuousInheritedStoryPreservesChangedEdgeDistance(bool footer) {
        using WordDocument source = CreateInheritedRunningStoryControl(footer, "default", "continuous");
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(18D);
        source.Sections[1].Margins.HeaderDistance = source.Sections[1].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(54D);
        source.Paragraphs.Single(p => p.Text == "S2A").PageBreakBefore = true;
        using var pdf = PdfPigDocument.Open(source.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        using var pdf = PdfPigDocument.Open(source.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
    public void SaveAsPdf_InheritedRequiredStoryPreservesMixedAnchoredTextBox(bool footer, bool even, bool table) {
        using WordDocument source = WordDocument.Create();
        WordHeaderFooter required = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        if (table) required.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "TABLE";
        else required.AddParagraph("COUNT ").AddField(WordFieldType.SectionPages);
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
        string[] before = AnchoredRunningStoryXml(document);
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        var page = pdf.GetPage(2);
        var prefix = AnchoredRunningLetters(page, "PREFIX");
        var suffix = AnchoredRunningLetters(page, "SUFFIX");
        AnchoredRunningLetters(page, "BOX");
        Assert.Equal(prefix[0].StartBaseLine.Y, suffix[0].StartBaseLine.Y, 3);
        Assert.Equal(prefix.Last().EndBaseLine.X, suffix[0].StartBaseLine.X, 3);
        Assert.DoesNotContain(page.Paths, path => path.IsStroked);
        Assert.Contains(table ? "TABLE" : "COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }
}
