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
    [InlineData(false, "color")]
    [InlineData(false, "alignment")]
    [InlineData(false, "font")]
    [InlineData(false, "family")]
    [InlineData(false, "size")]
    [InlineData(false, "offset")]
    [InlineData(true, "color")]
    [InlineData(true, "alignment")]
    [InlineData(true, "font")]
    [InlineData(true, "family")]
    [InlineData(true, "size")]
    [InlineData(true, "offset")]
    public void SaveAsPdf_EmptyRunningStoryRetainsFrameWithPresentationOverrides(bool footer, string setting) {
        Assert.Equal(Render(false), Render(true));

        (double Baseline, int Lines) Render(bool configure) {
            using WordDocument document = CreateEmptyRunningStoryDocument("implicit", footer);
            WordParagraph body = document.AddParagraph("A");
            body.LineSpacing = footer ? 364 : 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            body.KeepLinesTogether = false; body.AvoidWidowAndOrphan = false;
            if (footer) for (int index = 1; index < 40; index++) body.AddBreak().AddText("A");
            WordPdfExportOptions options = EmptyRunningOptions();
            if (configure) {
                options.PdfOptions = new PdfCreateOptions();
                if (footer) {
                    if (setting == "alignment") options.PdfOptions.FooterAlign = PdfAlign.Right;
                    else if (setting == "font") options.PdfOptions.FooterFont = PdfStandardFont.Courier;
                    else if (setting == "family") options.PdfOptions.FooterFontFamily = "Arial";
                    else if (setting == "size") options.PdfOptions.FooterFontSize = 24D;
                    else if (setting == "offset") options.PdfOptions.FooterOffsetY = 45D;
                    else options.PdfOptions.FooterTextColor = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0);
                } else {
                    if (setting == "alignment") options.PdfOptions.HeaderAlign = PdfAlign.Right;
                    else if (setting == "font") options.PdfOptions.HeaderFont = PdfStandardFont.Courier;
                    else if (setting == "family") options.PdfOptions.HeaderFontFamily = "Arial";
                    else if (setting == "size") options.PdfOptions.HeaderFontSize = 24D;
                    else if (setting == "offset") options.PdfOptions.HeaderOffsetY = 45D;
                    else options.PdfOptions.HeaderTextColor = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0);
                }
            }
            using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "A").ToArray();
            return (letters[0].StartBaseLine.Y, letters.Length);
        }
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void SaveAsPdf_ImplicitRunningVariantPreservesOtherStoryPresentation(bool footer, int emptyVariant) {
        using WordDocument document = CreateEmptyRunningStoryDocument("absent", footer);
        WordHeaderFooter[] stories = footer
            ? new WordHeaderFooter[] { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate }
            : new WordHeaderFooter[] { document.HeaderFirstOrCreate, document.HeaderEvenOrCreate, document.HeaderDefaultOrCreate };
        foreach (var child in stories[emptyVariant].ChildElements.ToArray()) child.Remove();
        for (int index = 0; index < stories.Length; index++) {
            if (index != emptyVariant) stories[index].AddParagraph("STORY" + index);
            WordParagraph body = document.AddParagraph("A");
            body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        WordPdfExportOptions options = EmptyRunningOptions();
        options.PdfOptions = new PdfCreateOptions();
        if (footer) options.PdfOptions.FooterOffsetY = 24D;
        else options.PdfOptions.HeaderOffsetY = 24D;
        using var withImplicit = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, withImplicit.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            if (index == emptyVariant) Assert.DoesNotContain("STORY", withImplicit.GetPage(index + 1).Text);
            else Assert.Contains("STORY" + index, withImplicit.GetPage(index + 1).Text);
        }
        // Remove only the implicit part's reference. Other variants must retain
        // their existing projected text and caller-configured positioning.
        W.HeaderFooterValues kind = emptyVariant == 0 ? W.HeaderFooterValues.First
            : emptyVariant == 1 ? W.HeaderFooterValues.Even : W.HeaderFooterValues.Default;
        foreach (var reference in document.Sections[0]._sectionProperties.ChildElements.Where(element =>
            footer ? element is W.FooterReference f && f.Type?.Value == kind
                : element is W.HeaderReference h && h.Type?.Value == kind).ToArray()) reference.Remove();
        using WordDocument withoutDocument = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using var withoutImplicit = PdfPigDocument.Open(withoutDocument.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, withoutImplicit.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            if (index == emptyVariant) {
                if (!footer) Assert.Equal(36D,
                    withoutImplicit.GetPage(index + 1).Letters.Single(letter => letter.Value == "A").StartBaseLine.Y -
                    withImplicit.GetPage(index + 1).Letters.Single(letter => letter.Value == "A").StartBaseLine.Y, 3);
                continue;
            }
            Assert.Equal(withoutImplicit.GetPage(index + 1).Letters.Select(letter => (letter.Value, letter.StartBaseLine)),
                withImplicit.GetPage(index + 1).Letters.Select(letter => (letter.Value, letter.StartBaseLine)));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_EmptyRunningStoryRegistersItsExclusiveDefaultStyleFont(bool footer) {
        (double Baseline, int Lines) absent = Render("absent");
        (double Baseline, int Lines) implicitStory = Render("implicit");
        (double Baseline, int Lines) explicitStory = Render("explicit");
        Assert.Equal(explicitStory, implicitStory);
        if (footer) {
            Assert.Equal(40, absent.Lines);
            Assert.Equal(37, implicitStory.Lines);
        } else {
            Assert.Equal(42D, absent.Baseline - implicitStory.Baseline, 3);
        }

        (double Baseline, int Lines) Render(string kind) {
            const string family = "OfficeIMO Tall Implicit Story";
            using WordDocument document = CreateEmptyRunningStoryDocument(kind, footer);
            W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
            normal.StyleRunProperties = new W.StyleRunProperties(
                new W.RunFonts { Ascii = family, HighAnsi = family }, new W.FontSize { Val = "48" });
            normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
                Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
            });
            styles.AppendChild(new W.Style(new W.StyleName { Val = "Independent Body" },
                new W.BasedOn { Val = "Normal" }, new W.StyleRunProperties(
                    new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "22" })) {
                Type = W.StyleValues.Paragraph, StyleId = "IndependentBody"
            });
            WordParagraph body = document.AddParagraph("A");
            body.LineSpacing = footer ? 364 : 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            body.KeepLinesTogether = false; body.AvoidWidowAndOrphan = false;
            body._paragraph.ParagraphProperties!.ParagraphStyleId = new W.ParagraphStyleId { Val = "IndependentBody" };
            if (footer) for (int index = 1; index < 40; index++) body.AddBreak().AddText("A");
            WordPdfExportOptions options = EmptyRunningOptions();
            options.PdfOptions = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                family, CreateFontWithLineMetrics(1200, -300, ' ')));
            using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "A").ToArray();
            return (letters[0].StartBaseLine.Y, letters.Length);
        }
    }

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
        body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        body._paragraph.ParagraphProperties!.ParagraphMarkRunProperties =
            new W.ParagraphMarkRunProperties(new W.Vanish { Val = false });

        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;

        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        body.KeepLinesTogether = false; body.AvoidWidowAndOrphan = false;
        for (int index = 1; index < 40; index++) body.AddBreak().AddText("A");
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.True(document.Sections[0].DifferentFirstPage);
        Assert.True(document.Sections[0].DocumentOddEvenSettingEnabled);
        Assert.Null(document.Header!.Default);
        Assert.NotNull(document.Header.First);
        Assert.NotNull(document.Header.Even);
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        double[] baselines = pdf.GetPages().Select(page =>
            Assert.Single(page.Letters, letter => letter.Value == "A").StartBaseLine.Y).ToArray();
        Assert.Equal(36D, baselines[2] - baselines[0], 3);
        Assert.Equal(baselines[0], baselines[1], 3);
    }

    private static double EmptyRunningHeaderBaseline(string kind, int distance) {
        using WordDocument document = CreateEmptyRunningStoryDocument(kind, footer: false);
        document.Sections[0].Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(((uint)distance) / 20D);
        WordParagraph body = document.AddParagraph("A");
        body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
        body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        string before = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.FirstOrDefault()?.Header.OuterXml ?? "";
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart!.HeaderParts.FirstOrDefault()?.Header.OuterXml ?? "");
        Assert.Equal(1, pdf.NumberOfPages);
        return Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A").StartBaseLine.Y;
    }

    private static WordDocument CreateEmptyRunningStoryDocument(string kind, bool footer) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
        source.Sections[0].Margins.Top = source.Sections[0].Margins.Bottom = global::OfficeIMO.OfficeLength.Points(30D);
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(36D);
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

    private static WordPdfExportOptions EmptyRunningOptions() => new() {
        IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
