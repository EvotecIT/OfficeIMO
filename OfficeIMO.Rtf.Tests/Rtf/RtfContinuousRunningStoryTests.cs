using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfContinuousRunningStoryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ContinuousDefaultHeaderKeepsCurrentPageUntilRealBreak(bool replaceHeader, bool bodyPageBreak) {
        // Native Word DOCX and reopened RTF controls independently establish
        // one/one/two pages and headers A / A / A then B for these transitions.
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("S1INTRO");
        first.AddParagraph("S1END");
        RtfSection second = source.AddSection(RtfSectionBreakKind.Continuous);
        if (replaceHeader) second.AddHeader().AddParagraph("HEADERB");
        second.AddParagraph("S2INTRO");
        second.AddParagraph("S2BEFOREBREAK");
        second.AddParagraph("S2FOLLOWING").PageBreakBefore = bodyPageBreak;
        second.AddParagraph("S2END");

        Converted converted = Convert(source, $"rtf-continuous-{replaceHeader}-{bodyPageBreak}");

        Assert.Equal(bodyPageBreak ? 2 : 1, converted.Pages.Length);
        Assert.Equal(1, Occurrences(converted.Pages[0], "HEADERA"));
        Assert.DoesNotContain("HEADERB", converted.Pages[0], StringComparison.Ordinal);
        if (bodyPageBreak) {
            Assert.Equal(1, Occurrences(converted.Pages[1], "HEADERB"));
            Assert.DoesNotContain("HEADERA", converted.Pages[1], StringComparison.Ordinal);
        }
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
        string[] body = { "S1INTRO", "S1END", "S2INTRO", "S2BEFOREBREAK", "S2FOLLOWING", "S2END" };
        for (int index = 0; index < body.Length; index++) {
            int expectedPage = bodyPageBreak && index >= 4 ? 1 : 0;
            Assert.Equal(1, Occurrences(converted.Pages[expectedPage], body[index]));
        }
        AssertBodyOnce(converted.Pages, body);
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void ContinuousHeaderRegistersItsSourceOnlyFontWithoutChangingBodyTypography() {
        RtfDocument source = CreateDocument(out RtfSection first);
        source.Settings.SetDefaultFont(source.AddFont("Arial"));
        int headerFont = source.AddFont("Courier New");
        first.AddParagraph("FIRSTBODY").Runs[0].FontSize = 10;
        RtfSection changed = source.AddSection(RtfSectionBreakKind.Continuous);
        RtfRun header = changed.AddHeader().AddParagraph().AddText("COURIERHEADER");
        header.FontId = headerFont;
        header.FontSize = 14;
        changed.AddParagraph("BEFOREBREAK").Runs[0].FontSize = 10;
        RtfParagraph following = changed.AddParagraph("FOLLOWINGPAGE");
        following.PageBreakBefore = true;
        following.Runs[0].FontSize = 10;

        Converted converted = Convert(source, "rtf-continuous-source-only-font");

        Assert.Equal(2, converted.Pages.Length);
        Assert.DoesNotContain("COURIERHEADER", converted.Pages[0], StringComparison.Ordinal);
        Assert.Equal(1, Occurrences(converted.Pages[1], "COURIERHEADER"));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(converted.Bytes);
        var paintedHeader = Assert.Single(independent.GetPage(2).GetWords(), word => word.Text == "COURIERHEADER");
        Assert.All(paintedHeader.Letters, letter => {
            Assert.Contains("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(14D, letter.PointSize, 2);
        });
        var body = Assert.Single(independent.GetPage(2).GetWords(), word => word.Text == "FOLLOWINGPAGE");
        Assert.All(body.Letters, letter => {
            Assert.DoesNotContain("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(10D, letter.PointSize, 2);
        });
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "BEFOREBREAK", "FOLLOWINGPAGE");
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Theory]
    [InlineData(@"\linex0")]
    [InlineData(@"\linemod0\linex360\linestarts7\linecont")]
    public void ImportedInactiveLineNumberSettingsDoNotBlockContinuousHeaders(string lineNumberSettings) {
        // Word writes a line-number distance even when numbering is inactive.
        // These syntax controls must survive while the visible header follows
        // the same physical-page ownership as an authored document model.
        string source = ImportedContinuousInput(lineNumberSettings,
            @"\pard\fs20\sl-240\slmult0 FIRSTBODY\par\sect", bodyPageBreak: true);
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(source);
        RtfReadResult read = RtfDocument.LoadResult(bytes);
        Assert.Equal(2, read.Document.Sections.Count);
        Assert.All(read.Document.Sections, section => {
            Assert.True(section.LineNumbering.HasAnyValue);
            Assert.Equal(0, section.LineNumbering.CountBy.GetValueOrDefault());
        });

        Converted converted = Convert(read.Document, "rtf-continuous-imported-inactive-numbering");

        Assert.Equal(2, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERB", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERA", converted.Pages[1], StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "BEFOREBREAK", "FOLLOWINGPAGE");
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
        Assert.Equal(bytes, read.ToBytesLossless());
    }

    [Fact]
    public void ImportedEmptySectionTerminatingParagraphKeepsItsLineAndInitialPageHeader() {
        // Native Word's own RTF terminates this authored empty paragraph
        // with a grouped section control, without an intervening par.
        string source = ImportedContinuousInput(@"\linex0",
            @"\pard\fs20\sl-240\slmult0 {\b0\sect}", bodyPageBreak: false);
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(source);
        RtfReadResult read = RtfDocument.LoadResult(bytes);
        RtfParagraph empty = Assert.IsType<RtfParagraph>(Assert.Single(read.Document.Sections[0].Blocks));
        Assert.Empty(empty.Inlines);
        Assert.Equal(-240, empty.LineSpacingTwips);
        Assert.False(empty.LineSpacingMultiple);

        Converted converted = Convert(read.Document, "rtf-continuous-imported-empty-section");

        string page = Assert.Single(converted.Pages);
        Assert.Contains("HEADERA", page, StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", page, StringComparison.Ordinal);
        Assert.Contains("FOOTERFIXED", page, StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "BEFOREBREAK", "FOLLOWINGPAGE");
        Assert.Equal(bytes, read.ToBytesLossless());
    }

    [Theory]
    [InlineData(@"\par\sect", 1, 12D)]
    [InlineData(@"\par\pard\sl-480\slmult0{\sect}", 1, 12D)]
    [InlineData(@"\par\par\sect", 2, 24D)]
    public void ImportedSectionBoundaryPreservesNativeParagraphAdvance(
        string boundary, int firstSectionParagraphs, double expectedAdvance) {
        // Independent native RTF exports establish that a section marker
        // does not add a second paragraph after par, even after pard. An
        // explicit additional par retains its own source-defined line.
        string source = ImportedContinuousInput(@"\linex0",
            @"\pard\fs20\sl-240\slmult0 FIRSTBODY" + boundary, bodyPageBreak: false);
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(source);
        RtfReadResult read = RtfDocument.LoadResult(bytes);
        Assert.Equal(firstSectionParagraphs, read.Document.Sections[0].Blocks.Count);

        Converted converted = Convert(read.Document, "rtf-continuous-imported-section-boundary");

        Assert.Single(converted.Pages);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "BEFOREBREAK", "FOLLOWINGPAGE");
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(converted.Bytes);
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = Assert.Single(words, word => word.Text == "FIRSTBODY").Letters[0].StartBaseLine.Y;
        double second = Assert.Single(words, word => word.Text == "BEFOREBREAK").Letters[0].StartBaseLine.Y;
        Assert.Equal(expectedAdvance, first - second, 3);
        Assert.Equal(bytes, read.ToBytesLossless());
    }

    [Fact]
    public void ImportedEmptyMiddleSectionRetainsItsParagraphWithoutPaintingItsPendingHeader() {
        // The native e04 control reserves one exact line for a genuinely
        // empty middle section. It keeps header A on the shared page.
        string source = ImportedContinuousInput(@"\linex0",
            @"\pard\fs20\sl-240\slmult0 FIRSTBODY\par\sect\sectd\sbknone\linex0" +
            @"{\header\pard\fs20 HEADERMID\par}\sect", bodyPageBreak: false);
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(source);
        RtfReadResult read = RtfDocument.LoadResult(bytes);
        Assert.Equal(3, read.Document.Sections.Count);
        RtfParagraph empty = Assert.IsType<RtfParagraph>(Assert.Single(read.Document.Sections[1].Blocks));
        Assert.Empty(empty.Inlines);
        Assert.Equal(-240, empty.LineSpacingTwips);

        Converted converted = Convert(read.Document, "rtf-continuous-imported-empty-middle-section");

        string page = Assert.Single(converted.Pages);
        Assert.Contains("HEADERA", page, StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERMID", page, StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", page, StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "BEFOREBREAK", "FOLLOWINGPAGE");
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(converted.Bytes);
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = Assert.Single(words, word => word.Text == "FIRSTBODY").Letters[0].StartBaseLine.Y;
        double second = Assert.Single(words, word => word.Text == "BEFOREBREAK").Letters[0].StartBaseLine.Y;
        Assert.Equal(24D, first - second, 3);
        Assert.Equal(bytes, read.ToBytesLossless());
    }

    [Fact]
    public void ContinuousHeaderReplacementWorksInsideALaterPageStartingSection() {
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("PREFIXBODY");
        RtfSection second = source.AddSection(RtfSectionBreakKind.NextPage);
        second.AddHeader().AddParagraph("HEADERC");
        second.AddParagraph("GROUPSTART");
        RtfSection continuous = source.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddHeader().AddParagraph("HEADERD");
        continuous.AddParagraph("CONTINUOUSBODY");
        continuous.AddParagraph("FOLLOWINGPAGE").PageBreakBefore = true;

        Converted converted = Convert(source, "rtf-continuous-nested-page-group");

        Assert.Equal(3, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERC", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERD", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("HEADERD", converted.Pages[2], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERC", converted.Pages[2], StringComparison.Ordinal);
        Assert.Contains("CONTINUOUSBODY", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("FOLLOWINGPAGE", converted.Pages[2], StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "PREFIXBODY", "GROUPSTART", "CONTINUOUSBODY", "FOLLOWINGPAGE");
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void ContinuousHeaderReplacementSurvivesAnInheritedSuccessor() {
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("FIRSTBODY");
        RtfSection changed = source.AddSection(RtfSectionBreakKind.Continuous);
        changed.AddHeader().AddParagraph("HEADERB");
        changed.AddParagraph("CHANGEDBODY");
        RtfSection inherited = source.AddSection(RtfSectionBreakKind.Continuous);
        inherited.AddParagraph("INHERITEDBODY");
        inherited.AddParagraph("FOLLOWINGPAGE").PageBreakBefore = true;

        Converted converted = Convert(source, "rtf-continuous-inherited-successor");

        Assert.Equal(2, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERB", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERA", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("INHERITEDBODY", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("FOLLOWINGPAGE", converted.Pages[1], StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "CHANGEDBODY", "INHERITEDBODY", "FOLLOWINGPAGE");
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void PageStartingSectionReplacesUnusedContinuousHeaderPresentation() {
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("FIRSTBODY");
        RtfSection continuous = source.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddHeader().AddParagraph("HEADERB");
        continuous.AddParagraph("CONTINUOUSBODY");
        RtfSection explicitPage = source.AddSection(RtfSectionBreakKind.NextPage);
        explicitPage.AddHeader().AddParagraph("HEADERC");
        explicitPage.AddParagraph("EXPLICITPAGEBODY");

        Converted converted = Convert(source, "rtf-continuous-explicit-page-owner");

        Assert.Equal(2, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("CONTINUOUSBODY", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERC", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("EXPLICITPAGEBODY", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", string.Concat(converted.Pages), StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "CONTINUOUSBODY", "EXPLICITPAGEBODY");
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
    }

    [Theory]
    [InlineData("footer", "Footers")]
    [InlineData("first", "StorySelection")]
    [InlineData("note", "Headers")]
    [InlineData("line-numbering", "Headers")]
    public void UnsupportedContinuousStoryReplacementRetainsItsWarning(string unsupported, string expectedProperty) {
        RtfDocument source = CreateDocument(out RtfSection first);
        RtfParagraph firstBody = first.AddParagraph("FIRSTBODY");
        RtfSection continuous = source.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddHeader().AddParagraph("HEADERB");
        continuous.AddParagraph("CONTINUOUSBODY");
        if (unsupported == "footer") continuous.AddFooter().AddParagraph("REPLACEMENTFOOTER");
        else if (unsupported == "first") {
            continuous.PageSetup.DifferentFirstPageHeaderFooter = true;
            continuous.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("FIRSTVARIANT");
        } else if (unsupported == "line-numbering") first.LineNumbering.Set(countBy: 2);
        else firstBody.AddFootnote("1", "NOTEBODY");

        Converted converted = Convert(source, "rtf-continuous-unsupported-" + unsupported);

        PdfCore.PdfConversionWarning warning = Assert.Single(converted.Result.Warnings,
            item => item.Code == "ContinuousSectionPageSettingsFlattened");
        Assert.Equal(nameof(RtfConversionAction.Flattened), warning.Details["RtfAction"]);
        Assert.Contains(expectedProperty, warning.Details["Properties"], StringComparison.Ordinal);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", string.Concat(converted.Pages), StringComparison.Ordinal);
        Assert.DoesNotContain("REPLACEMENTFOOTER", string.Concat(converted.Pages), StringComparison.Ordinal);
        Assert.DoesNotContain("FIRSTVARIANT", string.Concat(converted.Pages), StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "CONTINUOUSBODY");
        if (unsupported == "note") Assert.Equal(1, Occurrences(string.Concat(converted.Pages), "NOTEBODY"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AuthoredEmptySectionsMatchTheirSerializedPhysicalOwnershipAndLineAdvance(bool emptyMiddle, bool laterPageGroup) {
        RtfDocument source = CreateDocument(out RtfSection first);
        source.Settings.SetDefaultFont(source.AddFont("Arial"));
        var body = new List<string>();
        RtfSection pageOwner = first;
        if (laterPageGroup) {
            Body(first, "PREFIXBODY");
            body.Add("PREFIXBODY");
            pageOwner = source.AddSection(RtfSectionBreakKind.NextPage);
            pageOwner.AddHeader().AddParagraph("HEADERC");
        }

        RtfSection empty = pageOwner;
        if (emptyMiddle) {
            Body(pageOwner, "FIRSTBODY");
            body.Add("FIRSTBODY");
            empty = source.AddSection(RtfSectionBreakKind.Continuous);
            empty.AddHeader().AddParagraph("HEADERMID");
        }
        int emptyIndex = source.Sections.Count - 1;
        RtfSection changed = source.AddSection(RtfSectionBreakKind.Continuous);
        changed.AddHeader().AddParagraph("HEADERB");
        Body(changed, "BEFOREBREAK");
        Body(changed, "FOLLOWINGPAGE").PageBreakBefore = true;
        body.Add("BEFOREBREAK");
        body.Add("FOLLOWINGPAGE");
        foreach (RtfParagraph paragraph in source.HeaderFooters.SelectMany(story => story.Paragraphs)) {
            paragraph.SpaceBeforeTwips = 0;
            paragraph.SpaceAfterTwips = 0;
        }
        Assert.Empty(empty.Blocks);
        string serialized = source.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        byte[] original = System.Text.Encoding.ASCII.GetBytes(serialized);
        RtfReadResult read = RtfDocument.LoadResult(original);
        RtfParagraph implicitParagraph = Assert.IsType<RtfParagraph>(Assert.Single(read.Document.Sections[emptyIndex].Blocks));
        Assert.Empty(implicitParagraph.Inlines);

        string suffix = $"{emptyMiddle}-{laterPageGroup}";
        Converted direct = Convert(source, "rtf-authored-empty-direct-" + suffix);
        Converted reloaded = Convert(read.Document, "rtf-authored-empty-reloaded-" + suffix);

        int ownerPage = laterPageGroup ? 1 : 0;
        Assert.Equal(ownerPage + 2, direct.Pages.Length);
        Assert.Equal(direct.Pages.Length, reloaded.Pages.Length);
        Assert.Equal(direct.Pages, reloaded.Pages);
        foreach (Converted converted in new[] { direct, reloaded }) {
            Assert.Contains(laterPageGroup ? "HEADERC" : "HEADERA", converted.Pages[ownerPage], StringComparison.Ordinal);
            Assert.DoesNotContain("HEADERB", converted.Pages[ownerPage], StringComparison.Ordinal);
            Assert.Equal(1, Occurrences(converted.Pages[ownerPage + 1], "HEADERB"));
            Assert.DoesNotContain("HEADERMID", string.Concat(converted.Pages), StringComparison.Ordinal);
            Assert.Contains("BEFOREBREAK", converted.Pages[ownerPage], StringComparison.Ordinal);
            Assert.Contains("FOLLOWINGPAGE", converted.Pages[ownerPage + 1], StringComparison.Ordinal);
            if (laterPageGroup) Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
            AssertBodyOnce(converted.Pages, body.ToArray());
            Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
        }

        using var directPdf = UglyToad.PdfPig.PdfDocument.Open(direct.Bytes);
        using var reloadedPdf = UglyToad.PdfPig.PdfDocument.Open(reloaded.Bytes);
        // Compare the section's line advance against a fresh physical page.
        // This protects the implicit paragraph without asserting unrelated
        // absolute origins for authored and parsed paragraph defaults.
        double directAdvance = Baseline(directPdf, "FOLLOWINGPAGE") - Baseline(directPdf, "BEFOREBREAK");
        double reloadedAdvance = Baseline(reloadedPdf, "FOLLOWINGPAGE") - Baseline(reloadedPdf, "BEFOREBREAK");
        Assert.Equal(reloadedAdvance, directAdvance, 3);
        Assert.Empty(empty.Blocks);
        Assert.Equal(serialized, source.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        Assert.Equal(original, read.ToBytesLossless());

        static RtfParagraph Body(RtfSection section, string text) {
            RtfParagraph paragraph = section.AddParagraph(text);
            paragraph.SpaceBeforeTwips = 0;
            paragraph.SpaceAfterTwips = 0;
            paragraph.LineSpacingTwips = 240;
            paragraph.LineSpacingMultiple = true;
            return paragraph;
        }

        static double Baseline(UglyToad.PdfPig.PdfDocument pdf, string marker) =>
            Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == marker)
                .Letters[0].StartBaseLine.Y;
    }

    [Theory]
    [InlineData(false, -480, false, 0, 0)]
    [InlineData(true, -480, false, 0, 0)]
    [InlineData(false, 480, true, 0, 0)]
    [InlineData(true, -480, false, 120, 180)]
    public void AuthoredEmptyMiddleSectionRetainsActiveParagraphControls(bool laterPageGroup,
        int lineSpacing, bool lineSpacingMultiple, int spaceBefore, int spaceAfter) {
        RtfDocument source = CreateDocument(out RtfSection first);
        source.Settings.SetDefaultFont(source.AddFont("Arial"));
        var body = new List<string>();
        RtfSection pageOwner = first;
        if (laterPageGroup) {
            Body(first, "PREFIXBODY");
            body.Add("PREFIXBODY");
            pageOwner = source.AddSection(RtfSectionBreakKind.NextPage);
            pageOwner.AddHeader().AddParagraph("HEADERC");
        }
        RtfParagraph preceding = Body(pageOwner, "FIRSTBODY");
        preceding.LineSpacingTwips = lineSpacing;
        preceding.LineSpacingMultiple = lineSpacingMultiple;
        preceding.SpaceBeforeTwips = spaceBefore;
        preceding.SpaceAfterTwips = spaceAfter;
        body.Add("FIRSTBODY");
        RtfSection empty = source.AddSection(RtfSectionBreakKind.Continuous);
        empty.AddHeader().AddParagraph("HEADERMID");
        int emptyIndex = source.Sections.Count - 1;
        RtfSection changed = source.AddSection(RtfSectionBreakKind.Continuous);
        changed.AddHeader().AddParagraph("HEADERB");
        Body(changed, "BEFOREBREAK");
        Body(changed, "FOLLOWINGPAGE").PageBreakBefore = true;
        body.Add("BEFOREBREAK");
        body.Add("FOLLOWINGPAGE");
        foreach (RtfParagraph paragraph in source.HeaderFooters.SelectMany(story => story.Paragraphs)) {
            paragraph.SpaceBeforeTwips = 0;
            paragraph.SpaceAfterTwips = 0;
        }
        Assert.Empty(empty.Blocks);
        string serialized = source.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        byte[] original = System.Text.Encoding.ASCII.GetBytes(serialized);
        RtfReadResult read = RtfDocument.LoadResult(original);
        RtfParagraph implicitParagraph = Assert.IsType<RtfParagraph>(Assert.Single(read.Document.Sections[emptyIndex].Blocks));
        Assert.Empty(implicitParagraph.Inlines);
        Assert.Equal(lineSpacing, implicitParagraph.LineSpacingTwips);
        Assert.Equal(lineSpacingMultiple, implicitParagraph.LineSpacingMultiple);
        Assert.Equal(spaceBefore, implicitParagraph.SpaceBeforeTwips);
        Assert.Equal(spaceAfter, implicitParagraph.SpaceAfterTwips);

        string suffix = $"{laterPageGroup}-{lineSpacing}-{lineSpacingMultiple}-{spaceBefore}-{spaceAfter}";
        Converted direct = Convert(source, "rtf-authored-active-empty-direct-" + suffix);
        Converted reloaded = Convert(read.Document, "rtf-authored-active-empty-reloaded-" + suffix);
        int ownerPage = laterPageGroup ? 1 : 0;
        Assert.Equal(ownerPage + 2, direct.Pages.Length);
        Assert.Equal(direct.Pages, reloaded.Pages);
        foreach (Converted converted in new[] { direct, reloaded }) {
            Assert.Contains(laterPageGroup ? "HEADERC" : "HEADERA", converted.Pages[ownerPage], StringComparison.Ordinal);
            Assert.DoesNotContain("HEADERB", converted.Pages[ownerPage], StringComparison.Ordinal);
            Assert.Equal(1, Occurrences(converted.Pages[ownerPage + 1], "HEADERB"));
            Assert.DoesNotContain("HEADERMID", string.Concat(converted.Pages), StringComparison.Ordinal);
            Assert.Contains("BEFOREBREAK", converted.Pages[ownerPage], StringComparison.Ordinal);
            Assert.Contains("FOLLOWINGPAGE", converted.Pages[ownerPage + 1], StringComparison.Ordinal);
            AssertBodyOnce(converted.Pages, body.ToArray());
            Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
        }
        using var directPdf = UglyToad.PdfPig.PdfDocument.Open(direct.Bytes);
        using var reloadedPdf = UglyToad.PdfPig.PdfDocument.Open(reloaded.Bytes);
        double directAdvance = Baseline(directPdf, "FOLLOWINGPAGE") - Baseline(directPdf, "BEFOREBREAK");
        double reloadedAdvance = Baseline(reloadedPdf, "FOLLOWINGPAGE") - Baseline(reloadedPdf, "BEFOREBREAK");
        Assert.Equal(reloadedAdvance, directAdvance, 3);
        Assert.Empty(empty.Blocks);
        Assert.Equal(serialized, source.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        Assert.Equal(original, read.ToBytesLossless());

        static RtfParagraph Body(RtfSection section, string text) {
            RtfParagraph paragraph = section.AddParagraph(text);
            paragraph.SpaceBeforeTwips = 0;
            paragraph.SpaceAfterTwips = 0;
            paragraph.LineSpacingTwips = 240;
            paragraph.LineSpacingMultiple = true;
            return paragraph;
        }
        static double Baseline(UglyToad.PdfPig.PdfDocument pdf, string marker) =>
            Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == marker)
                .Letters[0].StartBaseLine.Y;
    }

    private static RtfDocument CreateDocument(out RtfSection first) {
        RtfDocument source = RtfDocument.Create();
        source.PageSetup.SetPaperSize(4800, 6400);
        source.PageSetup.SetMargins(leftTwips: 480, rightTwips: 480, topTwips: 640, bottomTwips: 640);
        source.PageSetup.HeaderDistanceTwips = 360;
        source.PageSetup.FooterDistanceTwips = 360;
        first = source.AddSection();
        first.AddHeader().AddParagraph("HEADERA");
        first.AddFooter().AddParagraph("FOOTERFIXED");
        return source;
    }

    private static string ImportedContinuousInput(string lineNumberSettings, string firstBody, bool bodyPageBreak) =>
        @"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Arial;}}" +
        @"\paperw4800\paperh6400\margl480\margr480\margt640\margb640\headery360\footery360" +
        @"\sectd" + lineNumberSettings + @"{\header\pard\fs20 HEADERA\par}{\footer\pard\fs20 FOOTERFIXED\par}" +
        firstBody + @"\sectd\sbknone" + lineNumberSettings +
        @"{\header\pard\fs20 HEADERB\par}\pard\fs20\sl-240\slmult0 BEFOREBREAK\par" +
        (bodyPageBreak ? @"\pagebb " : " ") + @"FOLLOWINGPAGE\par}";

    private static Converted Convert(RtfDocument source, string artifactName) {
        string before = source.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(new RtfToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        byte[] bytes = result.ToBytes();
        Assert.Equal(before, source.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        string[] pages = PdfCore.PdfReadDocument.Open(bytes,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).Pages
            .Select(page => new string(page.ExtractText().Where(character => !char.IsWhiteSpace(character)).ToArray())).ToArray();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(independent.NumberOfPages, pages.Length);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } directory) {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, artifactName + ".pdf"), bytes);
        }
        return new Converted(result, pages, bytes);
    }

    private static void AssertBodyOnce(string[] pages, params string[] markers) {
        string text = string.Concat(pages);
        int previous = -1;
        foreach (string marker in markers) {
            Assert.Equal(1, Occurrences(text, marker));
            int position = text.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(position > previous, "Body markers must retain their source order.");
            previous = position;
        }
    }

    private static int Occurrences(string text, string marker) {
        int count = 0;
        for (int index = 0; (index = text.IndexOf(marker, index, StringComparison.Ordinal)) >= 0; index += marker.Length) count++;
        return count;
    }

    private sealed class Converted {
        internal Converted(PdfCore.PdfDocumentConversionResult result, string[] pages, byte[] bytes) {
            Result = result;
            Pages = pages;
            Bytes = bytes;
        }
        internal PdfCore.PdfDocumentConversionResult Result { get; }
        internal string[] Pages { get; }
        internal byte[] Bytes { get; }
    }
}
