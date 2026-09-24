using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextSearchLineWrapTests {
    private const string WrappedContent =
        "BT /F1 12 Tf 50 700 Td (The quick brown needle) Tj 0 -14 Td (marker appears here) Tj ET\n";

    [Fact]
    public void DenseLineFragmentsStopAtTheConfiguredFlowComparisonBudget() {
        byte[] pdf = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (alpha) Tj ET\n" +
            "BT /F1 12 Tf 250 700 Td (beta) Tj ET\n" +
            "BT /F1 12 Tf 450 700 Td (gamma) Tj ET\n");
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxTextSearchFlowComparisons = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).Text.Find("alpha", readOptions: options));

        Assert.Equal(PdfReadLimitKind.TextSearchFlowComparisons, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void PhraseWrappedAcrossLinesIsOneHitWithBoundsForEachLine() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(WrappedContent));
        PdfTextMatch needle = Assert.Single(document.Text.Find("needle"));
        PdfTextMatch marker = Assert.Single(document.Text.Find("marker"));

        PdfTextMatch match = Assert.Single(document.Text.Find("needle marker"));

        Assert.Equal(1, match.PageNumber);
        Assert.Equal("needle marker", match.Text);
        Assert.Equal(2, match.VisualLineBounds.Count);
        AssertSameBounds(needle.VisualBounds, match.VisualLineBounds[0]);
        AssertSameBounds(marker.VisualBounds, match.VisualLineBounds[1]);
        Assert.True(match.VisualLineBounds[0].Bottom <= match.VisualLineBounds[1].Top + 0.5D);
        Assert.Equal(Math.Min(needle.VisualBounds.Left, marker.VisualBounds.Left), match.VisualBounds.Left, 3);
        Assert.Equal(needle.VisualBounds.Top, match.VisualBounds.Top, 3);
        Assert.Equal(Math.Max(needle.VisualBounds.Right, marker.VisualBounds.Right), match.VisualBounds.Right, 3);
        Assert.Equal(marker.VisualBounds.Bottom, match.VisualBounds.Bottom, 3);
        Assert.True(match.Y <= marker.Y + 0.001D && match.Y + match.Height >= needle.Y + needle.Height - 0.001D);
    }

    [Fact]
    public void SameLinePhraseKeepsSingleLineBounds() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(WrappedContent));

        PdfTextMatch match = Assert.Single(document.Text.Find("quick brown"));

        Assert.Equal("quick brown", match.Text);
        PdfSelectionQuad line = Assert.Single(match.VisualLineBounds);
        AssertSameBounds(match.VisualBounds, line);
    }

    [Fact]
    public void WrappedPhraseKeepsCaseAndWholeWordSemantics() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(WrappedContent));

        PdfTextMatch insensitive = Assert.Single(document.Text.Find("NEEDLE  MARKER"));
        Assert.Equal("needle marker", insensitive.Text);
        Assert.Empty(document.Text.Find("NEEDLE MARKER", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Single(document.Text.Find("needle marker", new PdfTextSearchOptions { MatchCase = true }));
        Assert.Single(document.Text.Find("needle marker", new PdfTextSearchOptions { WholeWords = true }));
        Assert.Empty(document.Text.Find("needle mark", new PdfTextSearchOptions { WholeWords = true }));
    }

    [Fact]
    public void LineEndHyphenationMatchesJoinedAndHyphenatedForms() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (Automatic hyph-) Tj 0 -14 Td (enation works) Tj ET\n"));

        PdfTextMatch joined = Assert.Single(document.Text.Find("hyphenation", new PdfTextSearchOptions { WholeWords = true }));
        PdfTextMatch hyphenated = Assert.Single(document.Text.Find("hyph-enation"));

        Assert.Equal("hyphenation", joined.Text);
        Assert.Equal(2, joined.VisualLineBounds.Count);
        Assert.Equal("hyph-enation", hyphenated.Text);
        Assert.Single(document.Text.Find("automatic hyphenation works"));
        Assert.Single(document.Text.Find("hyph-\nenation"));
        Assert.Single(document.Text.Find("hyph-\nenation", new PdfTextSearchOptions { WholeWords = true }));
        Assert.Single(document.Text.Find("hyph-\nenation", readOptions: new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxTextSearchMatches = 1 }
        }));
    }

    [Theory]
    [InlineData('\u00AD')]
    [InlineData('\u2010')]
    [InlineData('\u2011')]
    public void RetainedLineEndHyphensMatchOrdinaryHyphenatedQueries(char lineEndHyphen) {
        string source = "hyph" + lineEndHyphen + "\nenation";

        Assert.True(PdfTextSearchNormalization.Contains(source, "hyph-enation", StringComparison.Ordinal));
        Assert.True(PdfTextSearchNormalization.Contains(source, "hyphenation", StringComparison.Ordinal));
        Assert.True(PdfTextSearchNormalization.Contains(source, "hyph-\nenation", StringComparison.Ordinal));
        Assert.Equal((0, source.Length), Assert.Single(PdfTextSearchNormalization.FindSourceRanges(
            source, "hyph-enation", StringComparison.Ordinal)));
        Assert.False(PdfTextSearchNormalization.Contains("hyph" + lineEndHyphen + " next", "hyph-next", StringComparison.Ordinal));
    }

    [Fact]
    public void CombiningMarkBeforeLineEndHyphenMatchesBothWordForms() {
        const string source = "cafe\u0301-\nteria";

        Assert.True(PdfTextSearchNormalization.Contains(source, "cafe\u0301teria", StringComparison.Ordinal));
        Assert.True(PdfTextSearchNormalization.Contains(source, "cafe\u0301-teria", StringComparison.Ordinal));
        Assert.Equal((0, source.Length), Assert.Single(PdfTextSearchNormalization.FindSourceRanges(
            source, "cafe\u0301teria", StringComparison.Ordinal)));
        Assert.False(PdfTextSearchNormalization.Contains("\u0301-\nteria", "\u0301teria", StringComparison.Ordinal));
    }

    [Fact]
    public void ExplicitParagraphBreakDoesNotJoinAdjacentTextFlows() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (paragraph end) Tj 0 -12 Td 0 -12 Td (paragraph start) Tj ET\n"));

        Assert.Empty(document.Text.Find("end paragraph"));
        Assert.Empty(document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("end paragraph")).Areas);
        Assert.Single(document.Text.Find("paragraph end"));
        Assert.Single(document.Text.Find("paragraph start"));
    }

    [Fact]
    public void TableRowsDoNotFormOneWrappedSearchPhrase() {
        byte[] pdf = BuildRawTextPdf(
            "100 640 m 400 640 l 400 580 l 100 580 l h " +
            "250 640 m 250 580 l " +
            "100 620 m 400 620 l 100 600 m 400 600 l S\n" +
            "BT /F1 12 Tf 110 625 Td (Name) Tj ET\n" +
            "BT /F1 12 Tf 260 625 Td (Code) Tj ET\n" +
            "BT /F1 12 Tf 110 605 Td (Alpha) Tj ET\n" +
            "BT /F1 12 Tf 260 605 Td (100) Tj ET\n" +
            "BT /F1 12 Tf 110 585 Td (Beta) Tj ET\n" +
            "BT /F1 12 Tf 260 585 Td (200) Tj ET\n");

        Assert.NotEmpty(PdfReadDocument.Open(pdf).Pages[0].ExtractStructured().TablesDetailed);
        Assert.Empty(PdfDocument.Load(pdf).Text.Find("Alpha Beta"));
        Assert.Single(PdfDocument.Load(pdf).Text.Find("Alpha"));
        Assert.Empty(PdfDocument.Load(pdf).Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("Alpha Beta")).Areas);
    }

    [Fact]
    public void WrappedTextWithinOneTableCellRemainsSearchable() {
        byte[] pdf = BuildRawTextPdf(
            "100 660 m 400 660 l 400 560 l 100 560 l h " +
            "250 660 m 250 560 l " +
            "100 630 m 400 630 l 100 590 m 400 590 l S\n" +
            "BT /F1 12 Tf 110 640 Td (Name) Tj ET\n" +
            "BT /F1 12 Tf 260 640 Td (Code) Tj ET\n" +
            "BT /F1 12 Tf 110 610 Td (Alpha) Tj 0 -14 Td (Beta) Tj ET\n" +
            "BT /F1 12 Tf 260 610 Td (100) Tj ET\n" +
            "BT /F1 12 Tf 110 570 Td (Gamma) Tj ET\n" +
            "BT /F1 12 Tf 260 570 Td (200) Tj ET\n");
        Assert.NotEmpty(PdfReadDocument.Open(pdf).Pages[0].ExtractStructured().TablesDetailed);
        PdfDocument document = PdfDocument.Load(pdf);

        Assert.Single(document.Text.Find("Alpha Beta"));
        Assert.Empty(document.Text.Find("Beta Gamma"));
        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("Alpha Beta"));
        Assert.True(plan.Areas.Count >= 2);
        Assert.Empty(document.Redactions.Apply(plan).Text.Find("Alpha Beta"));
    }

    [Fact]
    public void TableDetectionHasItsOwnWorkBudget() {
        byte[] pdf = BuildRawTextPdf(WrappedContent);
        var limits = new PdfLoadOptions { Limits = new PdfReadLimits {
            MaxTextSearchFlowComparisons = 1,
            MaxTextSearchTableDetectionWork = 1
        } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).Text.Find("needle marker", readOptions: limits));
        Assert.Equal(PdfReadLimitKind.TextSearchTableDetectionWork, exception.Kind);
        Assert.Single(PdfDocument.Load(pdf).Text.Find("needle marker", readOptions: new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxTextSearchFlowComparisons = 1 }
        }));
    }

    [Fact]
    public void ActualTextLineEndHyphenationMatchesJoinedWord() {
        byte[] pdf = BuildRawTextPdf(
            "/Span << /ActualText <FEFF0068007900700068002D000A0065006E006100740069006F006E> >> BDC " +
            "BT /F1 12 Tf 50 700 Td (X) Tj ET EMC\n");
        PdfDocument document = PdfDocument.Load(pdf);

        Assert.Equal("hyph- enation", Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans()).Text);
        Assert.Single(document.Text.Find("hyphenation"));
        Assert.Single(document.Text.Find("hyph-enation"));
        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("hyphenation"));
        Assert.NotEmpty(plan.Areas);
        Assert.Empty(document.Redactions.Apply(plan).Text.Find("hyphenation"));
    }

    [Fact]
    public void PhrasesDoNotJoinIndependentColumns() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (left alpha) Tj 0 -14 Td (gamma) Tj ET\n" +
            "BT /F1 12 Tf 350 700 Td (beta right) Tj ET\n"));

        Assert.Empty(document.Text.Find("alpha beta"));
        Assert.Equal(2, Assert.Single(document.Text.Find("alpha gamma")).VisualLineBounds.Count);
        var redaction = new PdfRedactionSearchOptions();
        redaction.AddLiteral("alpha beta");
        Assert.Empty(document.Redactions.Search(redaction).Areas);
    }

    [Fact]
    public void RedactionFollowsWrappedListTextAcrossLogicalKinds() {
        byte[] pdf = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (- private needle) Tj 0 -14 Td (marker continues) Tj ET\n");
        PdfDocument document = PdfDocument.Load(pdf);
        PdfLogicalTextBlock[] blocks = PdfDocumentReadResult.From(PdfReadDocument.Open(pdf)).TextBlocks.ToArray();
        Assert.Contains(blocks, block => block.Kind == PdfLogicalElementKind.ListItem);
        Assert.Contains(blocks, block => block.Kind == PdfLogicalElementKind.TextBlock);
        Assert.Single(document.Text.Find("needle marker"));

        var search = new PdfRedactionSearchOptions().AddLiteral("needle marker");
        PdfRedactionPlan plan = document.Redactions.Search(search);

        Assert.True(plan.Areas.Count >= 2);
        Assert.Empty(document.Redactions.Apply(plan).Text.Find("needle marker"));
    }

    [Fact]
    public void RedactionFollowsEachColumnPastInterleavedBlocks() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (left alpha) Tj 0 -14 Td (gamma ends) Tj ET\n" +
            "BT /F1 12 Tf 350 700 Td (right beta) Tj 0 -14 Td (delta ends) Tj ET\n"));
        Assert.Single(document.Text.Find("alpha gamma"));
        Assert.Single(document.Text.Find("beta delta"));

        PdfRedactionPlan left = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("alpha gamma"));
        PdfRedactionPlan right = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("beta delta"));

        Assert.True(left.Areas.Count >= 2);
        Assert.True(right.Areas.Count >= 2);
        Assert.All(left.Areas, area => Assert.True(area.X < 300D));
        Assert.All(right.Areas, area => Assert.True(area.X > 300D));
    }

    [Fact]
    public void RedactionDoesNotJoinAcrossExcludedTableLine() {
        static PdfLogicalTextBlock Block(string text, double baseline, bool table = false) =>
            new(1, PdfLogicalElementKind.TextBlock, text, 50, 250, baseline, 12,
                new[] { new PdfTextSpan(text, "F1", 12, 50, baseline, 200) },
                isTableContent: table);

        PdfLogicalTextBlock[] blocks = {
            Block("private alpha", 700),
            Block("table value", 688, table: true),
            Block("beta public", 676)
        };
        Dictionary<int, string> matches = PdfRedactionPlanner.MatchLiteralsAcrossBlocks(blocks,
            new[] { "alpha beta" }, StringComparison.Ordinal, new PdfRedactionSearchWorkBudget("test"),
            _ => true, CancellationToken.None);

        Assert.Empty(matches);
    }

    [Fact]
    public void RedactionFlowBudgetAppliesOnlyToLiteralSearchAndUsesTheConfiguredLimit() {
        static PdfLogicalTextBlock Block(string text, double x) =>
            new(1, PdfLogicalElementKind.TextBlock, text, x, x + 80D, 700D, 12D,
                new[] { new PdfTextSpan(text, "F1", 12D, x, 700D, 80D) });
        PdfLogicalTextBlock[] blocks = { Block("alpha", 50D), Block("beta", 150D), Block("gamma", 250D) };
        var budget = new PdfRedactionSearchWorkBudget("test");

        Assert.Empty(PdfRedactionPlanner.MatchLiteralsAcrossBlocks(blocks, Array.Empty<string>(),
            StringComparison.Ordinal, budget, static _ => true, CancellationToken.None, maxFlowComparisons: 1));
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfRedactionPlanner.MatchLiteralsAcrossBlocks(blocks, new[] { "alpha beta" },
                StringComparison.Ordinal, budget, static _ => true, CancellationToken.None, maxFlowComparisons: 1));
        Assert.Equal(PdfReadLimitKind.TextSearchFlowComparisons, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void PublicRedactionSearchPassesTheConfiguredFlowBudgetOnlyForLiterals() {
        byte[] pdf = BuildRawTextPdf(
            "BT /F1 12 Tf 50 700 Td (alpha) Tj ET\n" +
            "BT /F1 12 Tf 250 700 Td (beta) Tj ET\n" +
            "BT /F1 12 Tf 450 700 Td (gamma) Tj ET\n");
        PdfDocument document = PdfDocument.Load(pdf);
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxTextSearchFlowComparisons = 1 }
        };

        Assert.NotEmpty(document.Redactions.Search(new PdfRedactionSearchOptions().AddRegex("alpha"),
            options: options).Areas);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("alpha beta"), options: options));
        Assert.Equal(PdfReadLimitKind.TextSearchFlowComparisons, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void RedactionFlowBudgetIsSharedAcrossNativeAndOcrBlocksOnOnePage() {
        static PdfLogicalTextBlock Block(string text, double x, PdfLogicalContentSourceKind sourceKind) =>
            new(1, PdfLogicalElementKind.TextBlock, text, x, x + 80D, 700D, 12D,
                new[] { new PdfTextSpan(text, "F1", 12D, x, 700D, 80D) }, sourceKind);
        PdfLogicalTextBlock[] blocks = {
            Block("native alpha", 50D, PdfLogicalContentSourceKind.Native),
            Block("native beta", 150D, PdfLogicalContentSourceKind.Native),
            Block("ocr alpha", 50D, PdfLogicalContentSourceKind.Ocr),
            Block("ocr beta", 150D, PdfLogicalContentSourceKind.Ocr)
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfRedactionPlanner.MatchLiteralsAcrossBlocks(blocks, new[] { "alpha beta" },
                StringComparison.Ordinal, new PdfRedactionSearchWorkBudget("test"), static _ => true,
                CancellationToken.None, maxFlowComparisons: 1));
        Assert.Equal(PdfReadLimitKind.TextSearchFlowComparisons, exception.Kind);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void RedactionFollowsRotatedWrappedLines() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(
            "BT /F1 12 Tf 0 1 -1 0 200 600 Tm (needle) Tj 0 -14 Td (marker) Tj ET\n"));
        Assert.Single(document.Text.Find("needle marker"));

        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("needle marker"));

        Assert.True(plan.Areas.Count >= 2);
        Assert.Empty(document.Redactions.Apply(plan).Text.Find("needle marker"));
    }

    [Fact]
    public void WrappedReplacementFitsAtTheFirstLineInsertionPoint() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(WrappedContent));
        PdfTextMatch match = Assert.Single(document.Text.Find("needle marker"));

        Assert.Throws<NotSupportedException>(() => document.Text.Replace(match, "long replacement",
            new PdfTextEditOptions { RegionWidthPolicy = PdfTextRegionWidthPolicy.RejectOverflow }));
        PdfTextEditResult result = document.Text.Replace(match, "long replacement", new PdfTextEditOptions {
            RegionWidthPolicy = PdfTextRegionWidthPolicy.ShrinkToFit, MinimumFontSize = 1D
        });
        Assert.True(Assert.Single(result.Document.Text.Find("long replacement")).FontSize < match.FontSize);
    }

    [Fact]
    public void GeneratedWrappedParagraphFindsPhraseAcrossTheWrap() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text(string.Join(" ", Enumerable.Range(1, 60).Select(index => "Needle marker " + index + " appears"))))
            .ToBytes();
        PdfTextSpan[] spans = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans()
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text)).ToArray();
        double[] baselines = spans.Select(static span => Math.Round(span.Y, 1)).Distinct().OrderByDescending(static y => y).ToArray();
        Assert.True(baselines.Length >= 2, "The generated paragraph must wrap.");
        string firstLine = string.Join(" ", spans.Where(span => Math.Round(span.Y, 1) == baselines[0]).OrderBy(static span => span.X).Select(static span => span.Text.Trim()));
        string secondLine = string.Join(" ", spans.Where(span => Math.Round(span.Y, 1) == baselines[1]).OrderBy(static span => span.X).Select(static span => span.Text.Trim()));
        string phrase = firstLine.Split(' ').Last() + " " + secondLine.Split(' ').First();

        IReadOnlyList<PdfTextMatch> matches = PdfDocument.Load(pdf).Text.Find(phrase);

        Assert.Contains(matches, static match => match.VisualLineBounds.Count == 2);
    }

    [Fact]
    public void ReplaceAllRewritesWrappedPhrase() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(WrappedContent));

        PdfTextEditResult result = document.Text.ReplaceAll("needle marker", "pin", new PdfTextSearchOptions { MatchCase = true });
        string text = result.Document.Reader.Text();

        Assert.Equal(1, result.AffectedCount);
        Assert.Empty(result.Document.Text.Find("needle"));
        Assert.Empty(result.Document.Text.Find("marker"));
        Assert.Contains("pin", text, StringComparison.Ordinal);
        Assert.Contains("appears here", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RedactionSearchFindsAndRemovesWrappedPhrase() {
        PdfDocument document = PdfDocument.Load(BuildRawTextPdf(WrappedContent));
        PdfTextMatch match = Assert.Single(document.Text.Find("needle marker"));
        var options = new PdfRedactionSearchOptions();
        options.AddLiteral("Needle Marker");

        PdfRedactionPlan plan = document.Redactions.Search(options);

        Assert.NotEmpty(plan.Areas);
        Assert.All(plan.Areas, static area => Assert.Equal(1, area.PageNumber));
        double bottom = plan.Areas.Min(static area => area.Y);
        double top = plan.Areas.Max(static area => area.Y + area.Height);
        // The two text lines sit on baselines 700 and 686; the planned areas must span both.
        Assert.True(bottom < 686D && top > 700D && match.Y < 686D && match.Y + match.Height > 700D,
            $"Areas {bottom}..{top} and match {match.Y}..{match.Y + match.Height} must span both baselines.");
        PdfDocument redacted = document.Redactions.Apply(plan);
        Assert.Empty(redacted.Text.Find("needle"));
        Assert.Empty(redacted.Text.Find("marker"));
    }

    private static void AssertSameBounds(PdfSelectionQuad expected, PdfSelectionQuad actual) {
        Assert.Equal(expected.Left, actual.Left, 3);
        Assert.Equal(expected.Top, actual.Top, 3);
        Assert.Equal(expected.Right, actual.Right, 3);
        Assert.Equal(expected.Bottom, actual.Bottom, 3);
    }

    private static byte[] BuildRawTextPdf(string content) {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
