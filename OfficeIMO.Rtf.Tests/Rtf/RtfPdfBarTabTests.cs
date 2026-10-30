using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfBarTabTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BarWithoutTabPaintsAParagraphRuleWithoutChangingSourceOrGlyphFlow(bool parsed) {
        RtfDocument control = Source("BAR", bar: false, positionTab: false, parsed: parsed);
        RtfDocument source = Source("BAR", bar: true, positionTab: false, parsed: parsed);
        string before = Snapshot(source);
        PdfCore.PdfDocumentConversionResult baseline = control.ConvertToPdfWithEvidence(Options());
        baseline.RequireNoLoss();

        PdfCore.PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(Options());
        byte[] bytes = result.ToBytes();

        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5), StringComparison.Ordinal);
        using var expected = PdfPigDocument.Open(baseline.ToBytes());
        using var actual = PdfPigDocument.Open(bytes);
        AssertSameGlyphFlow(expected, actual);
        var rule = Assert.Single(VerticalRules(actual.GetPage(1)));
        Assert.Equal(actual.GetPage(1).Letters[0].StartBaseLine.X + 36D, rule.Left, 3);
        Assert.True(rule.Height > 0D);
        AssertApproximation(result);
        Assert.Equal(before, Snapshot(source));
    }

    [Fact]
    public void BarDoesNotConsumeTheFollowingPositionalTabAdvance() {
        RtfDocument control = Source("A\tB", bar: false, positionTab: true, parsed: false);
        RtfDocument source = Source("A\tB", bar: true, positionTab: true, parsed: false);
        string before = Snapshot(source);
        PdfCore.PdfDocumentConversionResult baseline = control.ConvertToPdfWithEvidence(Options());
        baseline.RequireNoLoss();

        PdfCore.PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(Options());

        using var expected = PdfPigDocument.Open(baseline.ToBytes());
        using var actual = PdfPigDocument.Open(result.ToBytes());
        AssertSameGlyphFlow(expected, actual);
        var letters = actual.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var second = Assert.Single(letters, letter => letter.Value == "B");
        Assert.Equal(108D, second.StartBaseLine.X - first.StartBaseLine.X, 3);
        Assert.Equal(first.StartBaseLine.X + 36D, Assert.Single(VerticalRules(actual.GetPage(1))).Left, 3);
        AssertApproximation(result);
        Assert.Equal(before, Snapshot(source));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunningStoryBarPaintsUnlessExplicitPlainPresentationOmitsIt(bool plainPresentation) {
        RtfDocument source = Source("BODY", bar: false, positionTab: false, parsed: false);
        source.AddHeader().AddParagraph("HEADER").AddTabStop(720, RtfTabAlignment.Bar);
        string before = Snapshot(source);

        RtfToPdfOptions options = Options();
        if (plainPresentation) options.PdfOptions!.HeaderAlign = PdfCore.PdfAlign.Right;
        PdfCore.PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(options);
        using var actual = PdfPigDocument.Open(result.ToBytes());
        if (plainPresentation) {
            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "BarTabOmitted");
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.Equal("Omitted", warning.Details["RtfAction"]);
            Assert.DoesNotContain(result.Warnings, item => item.Code == "BarTabAppearanceApproximated");
            Assert.Empty(VerticalRules(actual.GetPage(1)));
        } else {
            AssertApproximation(result);
            Assert.DoesNotContain(result.Warnings, item => item.Code == "BarTabOmitted");
            var rule = Assert.Single(VerticalRules(actual.GetPage(1)));
            var header = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == "HEADER");
            Assert.Equal(header.Letters[0].StartBaseLine.X + 36D, rule.Left, 3);
        }
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        string text = PdfCore.PdfReadDocument.Open(result.ToBytes(),
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).ExtractText();
        Assert.Contains("BODY", text, StringComparison.Ordinal);
        Assert.Contains("HEADER", text, StringComparison.Ordinal);
        Assert.Equal(before, Snapshot(source));
    }

    [Fact]
    public void FlattenedNestedTableBarsReportOmissionAtEveryDepthWithoutChangingSource() {
        RtfDocument source = Source("BODY", bar: false, positionTab: false, parsed: false);
        source.AddStyle(7, "Inherited nested bar").AddTabStop(1440, RtfTabAlignment.Bar);
        RtfTableCell outer = source.AddTable(1, 1).Rows[0].Cells[0];
        outer.AddParagraph("OUTER");
        RtfTableCell nested = outer.AddTable(1, 1).Rows[0].Cells[0];
        nested.AddParagraph("NESTED").AddTabStop(720, RtfTabAlignment.Bar);
        nested.AddParagraph().AddTabStop(2160, RtfTabAlignment.Bar);
        nested.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("DEEP").SetStyle(7);
        string before = Snapshot(source);

        PdfCore.PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(Options());

        PdfCore.PdfConversionWarning[] omissions = result.Warnings.Where(item => item.Code == "BarTabOmitted").ToArray();
        Assert.Equal(3, omissions.Length);
        Assert.All(omissions, warning => {
            Assert.Equal("OfficeIMO.Rtf.Pdf", warning.Converter);
            Assert.Equal("TableCell/NestedTable/Paragraph/Tabs", warning.Source);
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.Equal("Omitted", warning.Details["RtfAction"]);
        });
        Assert.Equal(OfficeConversionLossKind.Approximation,
            Assert.Single(result.Warnings, item => item.Code == "NestedTableFlattened").LossKind);
        Assert.DoesNotContain(result.Warnings, item => item.Code == "BarTabAppearanceApproximated");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        byte[] bytes = result.ToBytes();
        using var actual = PdfPigDocument.Open(bytes);
        Assert.Empty(VerticalRules(actual.GetPage(1)));
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("BODY", text, StringComparison.Ordinal);
        Assert.Contains("OUTER", text, StringComparison.Ordinal);
        Assert.Contains("NESTED", text, StringComparison.Ordinal);
        Assert.Contains("DEEP", text, StringComparison.Ordinal);
        Assert.Equal(before, Snapshot(source));
    }

    private static RtfDocument Source(string text, bool bar, bool positionTab, bool parsed) {
        if (parsed) {
            string stops = (bar ? @"\tb720" : string.Empty) + (positionTab ? @"\tx2160" : string.Empty);
            return RtfDocument.Read(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Courier;}}\pard\f0\fs24\sl-400\slmult0" +
                stops + " " + text.Replace("\t", @"\tab ") + @"\par}").Document;
        }

        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Courier"));
        RtfParagraph paragraph = document.AddParagraph(text).SetLineSpacing(-400, multiple: false);
        if (bar) paragraph.AddTabStop(720, RtfTabAlignment.Bar);
        if (positionTab) paragraph.AddTabStop(2160, RtfTabAlignment.Left);
        return document;
    }

    // Built-in Courier keeps the contract proof independent of installed fonts.
    private static RtfToPdfOptions Options() => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        PdfOptions = new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Courier, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12),
            CompressContentStreams = false
        }
    };

    private static string Snapshot(RtfDocument document) => document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

    private static void AssertApproximation(PdfCore.PdfDocumentConversionResult result) {
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "BarTabAppearanceApproximated");
        Assert.Equal("OfficeIMO.Rtf.Pdf", warning.Converter);
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.Equal("Flattened", warning.Details["RtfAction"]);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
    }

    private static UglyToad.PdfPig.Core.PdfRectangle[] VerticalRules(UglyToad.PdfPig.Content.Page page) => page.Paths
        .Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
        .Where(bounds => bounds.HasValue && Math.Abs(bounds.Value.Width) < .001D && bounds.Value.Height > 1D)
        .Select(bounds => bounds!.Value).ToArray();

    private static void AssertSameGlyphFlow(PdfPigDocument expected, PdfPigDocument actual) {
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        for (int page = 1; page <= expected.NumberOfPages; page++) {
            var left = expected.GetPage(page).Letters;
            var right = actual.GetPage(page).Letters;
            Assert.Equal(left.Count, right.Count);
            for (int index = 0; index < left.Count; index++) {
                Assert.Equal(left[index].Value, right[index].Value);
                Assert.Equal(left[index].StartBaseLine.X, right[index].StartBaseLine.X, 3);
                Assert.Equal(left[index].StartBaseLine.Y, right[index].StartBaseLine.Y, 3);
                Assert.Equal(left[index].EndBaseLine.X, right[index].EndBaseLine.X, 3);
                Assert.Equal(left[index].EndBaseLine.Y, right[index].EndBaseLine.Y, 3);
            }
        }
    }
}
