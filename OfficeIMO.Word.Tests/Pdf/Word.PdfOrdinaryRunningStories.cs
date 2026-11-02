using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_OrdinaryRunningStoryAdmissionPreservesCollapsedSpacingAndHiddenMarks(bool hiddenMarks) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
        document.Sections[0].Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(18D);
        document.AddParagraph("BODY");
        WordHeaderFooter story = document.HeaderDefaultOrCreate;
        int count = hiddenMarks ? 100 : 4;
        for (int index = 0; index < count; index++) {
            WordParagraph paragraph = story.AddParagraph("ALPHA ");
            paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
                Before = "2000", After = "2000", Line = "240", LineRule = W.LineSpacingRuleValues.Exact });
            if (hiddenMarks && index + 1 < count) HideJoinMark(paragraph, true);
        }
        using var near = OpenJoinedParagraphPdf(document);
        document.Sections[0].Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(54D);
        using var far = OpenJoinedParagraphPdf(document);
        Assert.Equal(1, near.NumberOfPages);
        Assert.Equal(1, far.NumberOfPages);
        double[] baselines = near.GetPage(1).GetWords().Where(word => word.Text == "ALPHA")
            .Select(word => word.Letters[0].StartBaseLine.Y).Distinct().OrderByDescending(value => value).ToArray();
        Assert.True(baselines.Length > 1);
        Assert.InRange(baselines[0] - baselines[1], hiddenMarks ? 11.999D : 111.999D, hiddenMarks ? 12.001D : 112.001D);
        double nearY = near.GetPage(1).GetWords().First(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
        double farY = far.GetPage(1).GetWords().First(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
        Assert.InRange(farY - nearY, -36.001D, -35.999D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TallRunningFooterAllowsOneBodyLinePerPageWithinPhysicalMargins(bool nativeDoc) =>
        AssertTallRunningFooterAllowsOneBodyLine(nativeDoc, 48);

    [Theory]
    [InlineData(false, 32)]
    [InlineData(true, 32)]
    [InlineData(false, 40)]
    [InlineData(true, 40)]
    [InlineData(false, 64)]
    [InlineData(true, 64)]
    public void SaveAsPdf_TallRunningFooterPreservesHeaderAroundFirstLineFitBoundary(bool nativeDoc, int lineHeight) =>
        AssertTallRunningFooterAllowsOneBodyLine(nativeDoc, lineHeight);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TallRunningFooterPreservesHeaderWhenKeptBodyMustSplitAtPhysicalMargins(bool nativeDoc) =>
        AssertTallRunningFooterAllowsOneBodyLine(nativeDoc, 64, keepLinesTogether: true);

    private static void AssertTallRunningFooterAllowsOneBodyLine(bool nativeDoc, int lineHeight, bool keepLinesTogether = false) {
        using WordDocument source = CreateJoinedParagraphDocument();
        W.PageSize size = source.Sections[0]._sectionProperties.GetFirstChild<W.PageSize>()!;
        size.Width = 5200U; size.Height = 6000U;
        source.Sections[0].Margins.Top = global::OfficeIMO.OfficeLength.Points(72D);
        source.Sections[0].Margins.Bottom = global::OfficeIMO.OfficeLength.Points(72D);
        source.Sections[0].Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(36D);
        source.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(36D);
        foreach (WordHeaderFooter story in new WordHeaderFooter[] { source.HeaderDefaultOrCreate, source.FooterDefaultOrCreate })
            for (int index = 0; index < 3; index++)
                story.AddParagraph("STORY")._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
                    Before = "0", After = "0", Line = "640", LineRule = W.LineSpacingRuleValues.Exact });
        WordParagraph body = source.AddParagraph("ALPHA");
        body._paragraph.Append(new W.Run(new W.Break(), new W.Text("BETA")));
        body._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = (lineHeight * 20).ToString(), LineRule = W.LineSpacingRuleValues.Exact });
        if (keepLinesTogether) body.KeepLinesTogetherOverride = true;
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("ALPHA", pdf.GetPage(1).Text);
        Assert.DoesNotContain("BETA", pdf.GetPage(1).Text);
        Assert.Contains("BETA", pdf.GetPage(2).Text);
        Assert.Equal(pdf.GetPage(1).GetWords().Single(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y,
            pdf.GetPage(2).GetWords().Single(word => word.Text == "BETA").Letters[0].StartBaseLine.Y, 3);
        foreach (var page in pdf.GetPages()) {
            Assert.Equal(6, page.GetWords().Count(word => word.Text == "STORY"));
            var word = page.GetWords().Single(word => word.Text is "ALPHA" or "BETA");
            Assert.True(word.Letters[0].StartBaseLine.Y > 72D);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SaveAsPdf_OrdinaryRunningStoriesPreserveSourceDistanceAndExactParagraphSpacing(bool footer, bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        source.AddParagraph("BODY");
        WordHeaderFooter story = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        foreach (string text in new[] { "ALPHA", "BETA" })
            story.AddParagraph(text)._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
                Before = "120", After = "60", Line = "400", LineRule = W.LineSpacingRuleValues.Exact });
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        document.Sections[0].Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(18D);
        document.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(18D);
        string before = (footer ? document.Footer.Default : (WordHeaderFooter)document.Header.Default).ChildElements.First().Parent!.OuterXml;
        using var near = OpenJoinedParagraphPdf(document);
        document.Sections[0].Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(54D);
        document.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(54D);
        using var far = OpenJoinedParagraphPdf(document);
        foreach (var pdf in new[] { near, far }) {
            Assert.Equal(1, pdf.NumberOfPages);
            var words = pdf.GetPage(1).GetWords().ToArray();
            double alpha = Assert.Single(words, word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
            double beta = Assert.Single(words, word => word.Text == "BETA").Letters[0].StartBaseLine.Y;
            Assert.InRange(alpha - beta, 25.999D, 26.001D);
        }
        double nearY = near.GetPage(1).GetWords().Single(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
        double farY = far.GetPage(1).GetWords().Single(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
        Assert.InRange(farY - nearY, footer ? 35.999D : -36.001D, footer ? 36.001D : -35.999D);
        Assert.Equal(before, (footer ? document.Footer.Default : (WordHeaderFooter)document.Header.Default).ChildElements.First().Parent!.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_OrdinaryRunningStoryRetainsInitialBeforeSpacing(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.AddParagraph("BODY");
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph paragraph = story.AddParagraph("ALPHA");
        paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "400", LineRule = W.LineSpacingRuleValues.Exact });
        using var without = OpenJoinedParagraphPdf(document);
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines!.Before = "120";
        using var with = OpenJoinedParagraphPdf(document);
        double baselineWithout = without.GetPage(1).GetWords().Single(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
        double baselineWith = with.GetPage(1).GetWords().Single(word => word.Text == "ALPHA").Letters[0].StartBaseLine.Y;
        Assert.InRange(baselineWith - baselineWithout, footer ? -.001D : -6.001D, footer ? .001D : -5.999D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_OrdinaryRunningStoryVariantsUseExplicitCenterAndRightTabs(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.Sections[0]._sectionProperties.GetFirstChild<W.PageSize>()!.Width = 12240U;
        document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Points(72D);
        document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Points(72D);
        WordHeaderFooter[] stories = footer
            ? new WordHeaderFooter[] { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate }
            : new WordHeaderFooter[] { document.HeaderFirstOrCreate, document.HeaderEvenOrCreate, document.HeaderDefaultOrCreate };
        string[] labels = { "First", "Even", "Default" };
        for (int index = 0; index < stories.Length; index++) {
            WordParagraph paragraph = stories[index].AddParagraph(labels[index]);
            paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.Tabs(
                new W.TabStop { Val = W.TabStopValues.Center, Position = 4680 },
                new W.TabStop { Val = W.TabStopValues.Right, Position = 9360 }));
            paragraph._paragraph.Append(new W.Run(new W.TabChar(), new W.Text("Center"), new W.TabChar(), new W.Text("Right")));
        }
        for (int index = 0; index < 3; index++) {
            if (index > 0) document.AddParagraph().AddBreak(WordBreakType.Page);
            document.AddParagraph("BODY");
        }
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Equal(3, pdf.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            var words = pdf.GetPage(index + 1).GetWords().ToArray();
            var left = Assert.Single(words, word => word.Text == labels[index]);
            var center = Assert.Single(words, word => word.Text == "Center");
            var right = Assert.Single(words, word => word.Text == "Right");
            Assert.InRange(left.Letters[0].StartBaseLine.X, 71.999D, 72.001D);
            Assert.InRange((center.Letters[0].StartBaseLine.X + center.Letters.Last().EndBaseLine.X) / 2D, 305.999D, 306.001D);
            Assert.InRange(right.Letters.Last().EndBaseLine.X, 539.999D, 540.001D);
            Assert.Equal(left.Letters[0].StartBaseLine.Y, center.Letters[0].StartBaseLine.Y, 3);
            Assert.Equal(left.Letters[0].StartBaseLine.Y, right.Letters[0].StartBaseLine.Y, 3);
        }
    }
}
