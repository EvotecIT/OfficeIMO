using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("implicit", 720, 36D)]
    [InlineData("implicit", 0, 0D)]
    [InlineData("explicit", 720, 56D)]
    [InlineData("hidden", 720, 6D)]
    public void SaveAsPdf_EmptyRunningHeaderPreservesItsLineAndDistance(string kind, int distance, double expectedShift) {
        double withoutHeader = EmptyRunningHeaderBaseline("absent", distance);
        double withHeader = EmptyRunningHeaderBaseline(kind, distance);
        Assert.Equal(expectedShift, withoutHeader - withHeader, 3);
    }

    [Theory]
    [InlineData(false, 6D)]
    [InlineData(true, 36D)]
    public void SaveAsPdf_EmptyRunningHeaderInheritsDocumentMarkVisibility(bool visibleNormal, double expectedShift) {
        using WordDocument document = CreateEmptyRunningStoryDocument("implicit", footer: false);
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.AddChild(new W.Vanish(), true);
        if (visibleNormal) {
            W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
            normal.StyleRunProperties ??= new W.StyleRunProperties();
            normal.StyleRunProperties.AddChild(new W.Vanish { Val = false }, true);
        }
        WordParagraph body = document.AddParagraph("A");
        body.Hidden = false;
        body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
        body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;
        body._paragraph.ParagraphProperties!.ParagraphMarkRunProperties =
            new W.ParagraphMarkRunProperties(new W.Vanish { Val = false });

        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        double baseline = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A").StartBaseLine.Y;
        Assert.Equal(expectedShift, EmptyRunningHeaderBaseline("absent", 720) - baseline, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_EmptyRunningHeaderRetainsSpacingBefore(bool documentDefault) {
        using WordDocument document = CreateEmptyRunningStoryDocument("implicit", footer: false);
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        W.SpacingBetweenLines spacing = normal.StyleParagraphProperties!.GetFirstChild<W.SpacingBetweenLines>()!;
        if (documentDefault) {
            spacing.Before = null;
            W.ParagraphPropertiesBaseStyle defaults = styles.DocDefaults!.ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!;
            if (defaults.GetFirstChild<W.SpacingBetweenLines>() == null) defaults.AddChild(new W.SpacingBetweenLines(), true);
            spacing = defaults.GetFirstChild<W.SpacingBetweenLines>()!;
        }
        spacing.Before = "400";
        WordParagraph body = document.AddParagraph("A");
        body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
        body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;

        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        double baseline = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A").StartBaseLine.Y;
        Assert.Equal(56D, EmptyRunningHeaderBaseline("absent", 720) - baseline, 3);
    }

    [Theory]
    [InlineData("absent", 40)]
    [InlineData("implicit", 38)]
    [InlineData("explicit", 37)]
    [InlineData("hidden", 39)]
    public void SaveAsPdf_EmptyRunningFooterReservesItsBodyFrame(string kind, int expectedFirstPageLines) {
        using WordDocument document = CreateEmptyRunningStoryDocument(kind, footer: true);
        WordParagraph body = document.AddParagraph("A");
        body.LineSpacing = 364; body.LineSpacingRule = WordLineSpacingRule.Exact;
        body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;
        body.KeepLinesTogether = false; body.AvoidWidowAndOrphan = false;
        for (int index = 1; index < 40; index++) body.AddBreak().AddText("A");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(expectedFirstPageLines, pdf.GetPage(1).Letters.Count(letter => letter.Value == "A"));
        Assert.Equal(40, pdf.GetPages().Sum(page => page.Letters.Count(letter => letter.Value == "A")));
    }

    [Fact]
    public void SaveAsPdf_MissingRunningHeaderVariantDoesNotReserveAnEmptyStory() {
        using WordDocument source = CreateEmptyRunningStoryDocument("absent", footer: false);
        MainDocumentPart main = source._wordprocessingDocument.MainDocumentPart!;
        source.Sections[0]._sectionProperties.RemoveAllChildren<W.TitlePage>();
        source.Sections[0]._sectionProperties.AddChild(new W.TitlePage(), true);
        main.DocumentSettingsPart!.Settings.RemoveAllChildren<W.EvenAndOddHeaders>();
        main.DocumentSettingsPart!.Settings.AddChild(new W.EvenAndOddHeaders(), true);
        foreach (W.HeaderFooterValues variant in new[] { W.HeaderFooterValues.First, W.HeaderFooterValues.Even }) {
            HeaderPart part = main.AddNewPart<HeaderPart>();
            part.Header = new W.Header();
            source.Sections[0]._sectionProperties.PrependChild(new W.HeaderReference {
                Type = variant, Id = main.GetIdOfPart(part)
            });
        }
        for (int index = 0; index < 3; index++) {
            WordParagraph body = source.AddParagraph("A");
            body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
            body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.True(document.Sections[0].DifferentFirstPage);
        Assert.True(document.Sections[0].DocumentOddEvenSettingEnabled);
        Assert.Null(document.Header!.Default);
        Assert.NotNull(document.Header.First);
        Assert.NotNull(document.Header.Even);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(3, pdf.NumberOfPages);
        double[] baselines = pdf.GetPages().Select(page =>
            Assert.Single(page.Letters, letter => letter.Value == "A").StartBaseLine.Y).ToArray();
        Assert.Equal(36D, baselines[2] - baselines[0], 3);
        Assert.Equal(baselines[0], baselines[1], 3);
    }

    private static double EmptyRunningHeaderBaseline(string kind, int distance) {
        using WordDocument document = CreateEmptyRunningStoryDocument(kind, footer: false);
        document.Sections[0].Margins.HeaderDistance = (uint)distance;
        WordParagraph body = document.AddParagraph("A");
        body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
        body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;
        string before = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.FirstOrDefault()?.Header.OuterXml ?? "";
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart!.HeaderParts.FirstOrDefault()?.Header.OuterXml ?? "");
        Assert.Equal(1, pdf.NumberOfPages);
        return Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A").StartBaseLine.Y;
    }

    private static WordDocument CreateEmptyRunningStoryDocument(string kind, bool footer) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
        source.Sections[0].Margins.Top = source.Sections[0].Margins.Bottom = 600;
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = 720;
        W.Style normal = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "400", Line = "600", LineRule = W.LineSpacingRuleValues.Exact
        });
        if (kind != "absent") {
            W.Paragraph? paragraph = kind == "implicit" ? null : new W.Paragraph();
            if (kind == "hidden") paragraph!.ParagraphProperties = new W.ParagraphProperties(
                new W.ParagraphMarkRunProperties(new W.Vanish()));
            MainDocumentPart main = source._wordprocessingDocument.MainDocumentPart!;
            if (footer) {
                FooterPart part = main.AddNewPart<FooterPart>();
                part.Footer = paragraph == null ? new W.Footer() : new W.Footer(paragraph);
                source.Sections[0]._sectionProperties.AddChild(new W.FooterReference {
                    Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(part)
                }, true);
            } else {
                HeaderPart part = main.AddNewPart<HeaderPart>();
                part.Header = paragraph == null ? new W.Header() : new W.Header(paragraph);
                source.Sections[0]._sectionProperties.AddChild(new W.HeaderReference {
                    Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(part)
                }, true);
            }
        }
        return WordDocument.Load(new MemoryStream(source.ToBytes()));
    }

    private static WordToPdfOptions EmptyRunningOptions() => new() {
        IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
