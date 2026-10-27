using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInlineRuleTests {
    private static readonly PdfColor RuleColor = new(.2D, .5D, .8D);
    private static readonly PdfColor AdditionalColor = new(.8D, .2D, .4D);

    [Theory]
    [InlineData(100D, 40D)]
    [InlineData(24D, 24D)]
    public void CappedRuleClampsToTheActualFrame(double frameWidth, double expectedWidth) {
        var rule = new PdfInlineRule(.5D, RuleColor, maximumWidth: 40D, baselineOffset: 3D);
        byte[] bytes = Create(frameWidth).Paragraph(p => p.Runs(new[] { PdfTextRun.Inline(rule) }), style: Style()).ToBytes();
        var band = Assert.Single(ReadBands(bytes));
        Assert.Equal(20D, band.X, 3);
        Assert.Equal(expectedWidth, band.Width, 3);
        Assert.Equal(.5D, band.Height, 3);
    }

    [Fact]
    public void FillRuleConsumesOnlyTheWidthAfterLeadingTextAndWrapsFollowingText() {
        byte[] bytes = Create(100D).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("LEAD"), PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor)), PdfTextRun.Normal("TAIL")
        }), style: Style()).ToBytes();
        var band = Assert.Single(ReadBands(bytes));
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        var lead = letters.Take(4).ToArray();
        var tail = letters.Skip(4).ToArray();
        Assert.Equal("LEADTAIL", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(44D, band.X, 3);
        Assert.Equal(76D, band.Width, 3);
        Assert.Equal(lead[0].StartBaseLine.Y, lead[3].StartBaseLine.Y, 3);
        Assert.True(tail[0].StartBaseLine.Y < lead[0].StartBaseLine.Y);
        Assert.Equal(20D, tail[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(PdfAlign.Left, 20D)]
    [InlineData(PdfAlign.Center, 50D)]
    [InlineData(PdfAlign.Right, 80D)]
    public void ParagraphAlignmentUsesTheResolvedBandWidth(PdfAlign alignment, double expectedX) {
        byte[] bytes = Create(100D).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor, maximumWidth: 40D))
        }), alignment, style: Style()).ToBytes();
        var band = Assert.Single(ReadBands(bytes));
        Assert.Equal(expectedX, band.X, 3);
        Assert.Equal(40D, band.Width, 3);
    }

    [Fact]
    public void ReusingTheSameRuleAcrossFramesAndSerializationsDoesNotFreezeItsWidth() {
        var rule = new PdfInlineRule(.5D, RuleColor, maximumWidth: 40D);
        var run = PdfTextRun.Inline(rule);
        var narrow = Create(24D).Paragraph(p => p.Runs(new[] { run }), style: Style());
        var wide = Create(100D).Paragraph(p => p.Runs(new[] { run }), style: Style());
        Assert.Equal(24D, Assert.Single(ReadBands(narrow.ToBytes())).Width, 3);
        Assert.Equal(40D, Assert.Single(ReadBands(wide.ToBytes())).Width, 3);
        Assert.Equal(24D, Assert.Single(ReadBands(narrow.ToBytes())).Width, 3);
        Assert.Equal(0D, rule.Width);
        Assert.Same(rule, run.InlineElement);
    }

    [Fact]
    public void AdditionalBandRetainsItsColorAndBaselineWhileSharingTheResolvedWidth() {
        var original = new PdfInlineRule(.5D, RuleColor, maximumWidth: 40D, baselineOffset: 3D);
        var decorated = original.WithAdditionalBand(.75D, AdditionalColor, baselineOffset: -4D);
        PdfDocument Render(PdfInlineRule rule) => Create(60D).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("A"), PdfTextRun.Inline(rule), PdfTextRun.Normal("\nB")
        }), style: Style());

        byte[] originalBytes = Render(original).ToBytes();
        var document = Render(decorated);
        byte[] bytes = document.ToBytes();
        var primary = Assert.Single(ReadBands(bytes));
        var additional = Assert.Single(ReadBands(bytes, AdditionalColor));
        Assert.Equal(26D, primary.X, 3);
        Assert.Equal(40D, primary.Width, 3);
        Assert.Equal(primary.X, additional.X, 3);
        Assert.Equal(primary.Width, additional.Width, 3);
        Assert.Equal(-7D, additional.Y - primary.Y, 3);
        Assert.Equal(.75D, additional.Height, 3);
        Assert.Empty(ReadBands(originalBytes, AdditionalColor));
        Assert.Equal(Assert.Single(ReadBands(originalBytes)), primary);
        Assert.Equal(primary, Assert.Single(ReadBands(document.ToBytes())));
        Assert.Equal(additional, Assert.Single(ReadBands(document.ToBytes(), AdditionalColor)));
        using var reference = PdfPigDocument.Open(originalBytes);
        using var actual = PdfPigDocument.Open(bytes);
        Assert.Equal(Assert.Single(reference.GetPage(1).Letters, letter => letter.Value == "B").StartBaseLine.Y,
            Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Fact]
    public void ZeroRemainingWidthSuppressesEveryBandWithoutChangingTheSourceFontLineBox() {
        var original = new PdfInlineRule(.5D, RuleColor, baselineOffset: 3D);
        var decorated = original.WithAdditionalBand(.75D, AdditionalColor, baselineOffset: -40D);
        byte[] Render(PdfInlineRule rule) => Create(60D).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("0123456789"),
            PdfTextRun.Inline(rule, PdfTextRun.Normal(string.Empty, fontSize: 24D)),
            PdfTextRun.Normal("\nB")
        }), style: Style()).ToBytes();

        byte[] bytes = Render(decorated);
        Assert.Empty(ReadBands(bytes));
        Assert.Empty(ReadBands(bytes, AdditionalColor));
        using var reference = PdfPigDocument.Open(Render(original));
        using var actual = PdfPigDocument.Open(bytes);
        Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
        Assert.Equal(Assert.Single(reference.GetPage(1).Letters, letter => letter.Value == "B").StartBaseLine.Y,
            Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(PdfStandardFont.Helvetica, false, 120D, 1)]
    [InlineData(PdfStandardFont.TimesRoman, true, 120D, 1)]
    [InlineData(PdfStandardFont.Helvetica, false, 65D, 2)]
    public void PublicInlineBuilderRetainsSelectedTypographyForFollowingLinesAndPages(
        PdfStandardFont font, bool boldItalic, double pageHeight, int expectedPages) {
        byte[] reference = RenderBuilderFontMetrics(rule: false, font, boldItalic, pageHeight);
        byte[] actual = RenderBuilderFontMetrics(rule: true, font, boldItalic, pageHeight);
        AssertBuilderFollowingLineMatches(reference, actual, expectedPages);
    }

    [Fact]
    public void PublicInlineBuilderRetainsRegisteredNamedFontLineMetrics() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(
            1000, -400, 150, 400, 'R');
        var family = new PdfEmbeddedFontFamily("Builder rule metrics", font);
        byte[] reference = RenderBuilderFontMetrics(rule: false, PdfStandardFont.Helvetica, false, 120D, family);
        byte[] actual = RenderBuilderFontMetrics(rule: true, PdfStandardFont.Helvetica, false, 120D, family);
        AssertBuilderFollowingLineMatches(reference, actual, expectedPages: 1);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("table")]
    public void ThinRuleRetainsItsSourceRunFontLineBox(string frame) {
        using var reference = PdfPigDocument.Open(RenderFontMetrics(frame, rule: false));
        byte[] bytes = RenderFontMetrics(frame, rule: true);
        using var actual = PdfPigDocument.Open(bytes);
        var expectedNext = Assert.Single(reference.GetPage(1).Letters, letter => letter.Value == "B");
        var actualNext = Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "B");
        Assert.Equal(expectedNext.StartBaseLine.Y, actualNext.StartBaseLine.Y, 3);
        Assert.Equal(.5D, Assert.Single(ReadBands(bytes)).Height, 3);
    }

    [Fact]
    public void ZeroRemainingWidthRetainsSourceMetricsAndReferenceWithoutPaintingInk() {
        var document = Create(60D);
        var note = document.Content.DefineFootnote(body => body.Text("Note", style: Style()));
        var source = PdfTextRun.Normal(string.Empty).WithFootnoteReference(note);
        document.Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("0123456789"),
            PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor), source),
            PdfTextRun.Normal("\nB")
        }), style: Style());
        byte[] bytes = document.ToBytes();
        using var actual = PdfPigDocument.Open(bytes);
        using var reference = PdfPigDocument.Open(Create(60D).Paragraph(p => p.Text("0123456789\nB"), style: Style()).ToBytes());
        Assert.Empty(ReadBands(bytes));
        Assert.Equal(1, actual.NumberOfPages);
        Assert.Contains("Note", actual.GetPage(1).Text, StringComparison.Ordinal);
        var actualNext = Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "B");
        var expectedNext = Assert.Single(reference.GetPage(1).Letters, letter => letter.Value == "B");
        Assert.Equal(expectedNext.StartBaseLine.Y, actualNext.StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(PdfTabAlignment.Center, 68D, 52D)]
    [InlineData(PdfTabAlignment.Right, 56D, 64D)]
    public void AlignedTabLookaheadMeasuresIntrinsicContentBeforeResolvingTheRule(
        PdfTabAlignment alignment, double expectedX, double expectedWidth) {
        var style = Style();
        style.TabStops.Add(new PdfTabStop(60D, alignment));
        byte[] bytes = Create(100D).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("A"), PdfTextRun.Tab(),
            PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor)), PdfTextRun.Normal("TAIL")
        }), style: style).ToBytes();
        var band = Assert.Single(ReadBands(bytes));
        Assert.Equal(expectedX, band.X, 3);
        Assert.Equal(expectedWidth, band.Width, 3);
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        // The authored TAB positions the rule without adding a SPACE glyph.
        Assert.Equal("ATAIL", string.Concat(letters.Select(letter => letter.Value)));
        var lead = letters[0];
        Assert.Equal("A", lead.Value);
        var tail = Assert.Single(letters, letter => letter.Value == "T");
        Assert.True(tail.StartBaseLine.Y < lead.StartBaseLine.Y);
        Assert.Equal(20D, tail.StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutomaticTableAndRowWidthsDoNotExpandFlexibleRulesDuringIntrinsicMeasurement(bool importedGrid) {
        PdfTextRun[] tableRuns = {
            PdfTextRun.Normal("A\t"), PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor)), PdfTextRun.Normal("B")
        };
        var document = Create(100D);
        document.Table(new[] { new[] { PdfTableCell.RichTextCell(tableRuns) } }, style: new PdfTableStyle {
            FontSize = 10D, HeaderRowCount = 0, HeaderBold = false, AutoFitColumns = true,
            MaxWidth = 100D, CellPaddingX = 0D, CellPaddingY = 0D, BorderColor = null,
            ShrinkTextToFit = true, MinimumShrinkFontSize = 10D, AutoFitWidthUsesContentMinimum = importedGrid
        });
        document.Content.Row(row => row
            .AutoColumn(column => column.Paragraph(p => p.Runs(new[] {
                PdfTextRun.Normal("C"), PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor))
            }), style: Style()))
            .RelativeColumn(column => column.Text("D", style: Style())));
        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("ABCD", string.Concat(letters.Select(letter => letter.Value)));
        Assert.True(Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y <
            Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y);
        Assert.All(letters, letter => Assert.InRange(letter.StartBaseLine.X, 20D, 120D));
        var band = Assert.Single(ReadBands(bytes));
        Assert.True(band.Width > 0D);
        Assert.True(band.X + band.Width <= 120.001D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoWrapTableCellsStillResolveRulesAgainstTheirPhysicalFrame(bool paragraphCell) {
        PdfTextRun[] runs = { PdfTextRun.Normal("A"), PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor)) };
        var cell = (paragraphCell
            ? new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 10D) })
            : PdfTableCell.RichTextCell(runs)).WithNoWrap();
        byte[] bytes = Create(100D).Table(new[] { new[] { cell } }, style: new PdfTableStyle {
            FontSize = 10D, HeaderRowCount = 0, HeaderBold = false, CellPaddingX = 0D, CellPaddingY = 0D,
            ColumnWidthPoints = new() { 60D }, MaxWidth = 60D, PreserveWidth = true, BorderColor = null
        }).ToBytes();
        var band = Assert.Single(ReadBands(bytes));
        Assert.Equal(26D, band.X, 3);
        Assert.Equal(54D, band.Width, 3);
    }

    [Fact]
    public void PageContinuationRetainsTheFlexibleSourceAndItsLargeFontLineBox() {
        byte[] Render(bool rule) {
            var source = PdfTextRun.Normal("R", fontSize: 24D, font: PdfStandardFont.Helvetica);
            var document = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = 140D, PageHeight = 80D, MarginLeft = 20D, MarginRight = 20D,
                MarginTop = 20D, MarginBottom = 20D, DefaultFontSize = 8D,
                DefaultFont = PdfStandardFont.Courier, CompressContentStreams = false
            });
            return document.Paragraph(p => p.Runs(new[] {
                PdfTextRun.Normal("A\n", fontSize: 8D),
                rule ? PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor)
                    .WithAdditionalBand(.75D, AdditionalColor, baselineOffset: -2D), source) : source,
                PdfTextRun.Normal("\nB", fontSize: 8D)
            }), style: new PdfParagraphStyle {
                FontSize = 8D, LineSpacing = PdfLineSpacing.Multiple(1.4D), SpacingBefore = 0D, SpacingAfter = 0D
            }).ToBytes();
        }

        byte[] bytes = Render(rule: true);
        using var actual = PdfPigDocument.Open(bytes);
        using var reference = PdfPigDocument.Open(Render(rule: false));
        Assert.True(reference.NumberOfPages > 1);
        Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
        var expectedPage = Assert.Single(reference.GetPages(), page => page.Letters.Any(letter => letter.Value == "B"));
        var actualPage = Assert.Single(actual.GetPages(), page => page.Letters.Any(letter => letter.Value == "B"));
        Assert.Equal(expectedPage.Number, actualPage.Number);
        Assert.Equal(Assert.Single(expectedPage.Letters, letter => letter.Value == "B").StartBaseLine.Y,
            Assert.Single(actualPage.Letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
        var primary = Assert.Single(ReadBands(bytes));
        var additional = Assert.Single(ReadBands(bytes, AdditionalColor));
        Assert.Equal(100D, primary.Width, 3);
        Assert.Equal(primary.Width, additional.Width, 3);
        Assert.Equal(primary.X, additional.X, 3);
        Assert.Equal(-2D, additional.Y - primary.Y, 3);
    }

    private static PdfDocument Create(double frameWidth) => PdfDocument.Create(new PdfCreateOptions {
        PageWidth = frameWidth + 40D, PageHeight = 240D, MarginLeft = 20D, MarginRight = 20D,
        MarginTop = 20D, MarginBottom = 20D, DefaultFont = PdfStandardFont.Courier,
        DefaultFontSize = 10D, CompressContentStreams = false
    });

    private static PdfParagraphStyle Style() => new() {
        FontSize = 10D, LineSpacing = PdfLineSpacing.Multiple(1.2D), SpacingBefore = 0D, SpacingAfter = 0D
    };

    private static byte[] RenderBuilderFontMetrics(bool rule, PdfStandardFont font, bool boldItalic,
        double pageHeight, PdfEmbeddedFontFamily? family = null) {
        var options = new PdfCreateOptions {
            PageWidth = 200D, PageHeight = pageHeight, MarginLeft = 10D, MarginRight = 10D,
            MarginTop = 10D, MarginBottom = 10D, CompressContentStreams = false
        };
        if (family is not null) options.RegisterNamedFontFamily(family);
        var document = PdfDocument.Create(options);
        document.Content.Paragraph(p => {
            p.FontSize(24D).Bold(boldItalic).Italic(boldItalic);
            if (family is null) p.Font(font);
            else p.FontFamily(family.FamilyName);
            if (rule) p.Inline(new PdfInlineRule(.5D, RuleColor));
            else p.Text("R");
            p.ResetFontSize().ResetFont().Bold(false).Italic(false).Text("\nB");
        });
        return document.ToBytes();
    }

    private static void AssertBuilderFollowingLineMatches(byte[] referenceBytes, byte[] actualBytes, int expectedPages) {
        using var reference = PdfPigDocument.Open(referenceBytes);
        using var actual = PdfPigDocument.Open(actualBytes);
        Assert.Equal(expectedPages, reference.NumberOfPages);
        Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
        var expectedPage = Assert.Single(reference.GetPages(), page => page.Letters.Any(letter => letter.Value == "B"));
        var actualPage = Assert.Single(actual.GetPages(), page => page.Letters.Any(letter => letter.Value == "B"));
        Assert.Equal(expectedPage.Number, actualPage.Number);
        Assert.Equal(Assert.Single(expectedPage.Letters, letter => letter.Value == "B").StartBaseLine.Y,
            Assert.Single(actualPage.Letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
        Assert.Equal("B", string.Concat(actual.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value)));
        var band = Assert.Single(ReadBands(actualBytes));
        Assert.Equal(180D, band.Width, 3);
        Assert.Equal(.5D, band.Height, 3);
    }

    private static byte[] RenderFontMetrics(string frame, bool rule) {
        var source = PdfTextRun.Normal(rule ? string.Empty : "R", fontSize: 24D, font: PdfStandardFont.Helvetica);
        source.IsLineBoxStrut = rule;
        PdfTextRun[] runs = { rule ? PdfTextRun.Inline(new PdfInlineRule(.5D, RuleColor, 40D, 3D), source) : source,
            PdfTextRun.Normal("\nB", fontSize: 8D) };
        var document = Create(100D);
        var style = new PdfParagraphStyle { FontSize = 8D, LineSpacing = PdfLineSpacing.Multiple(1.4D), SpacingBefore = 0D, SpacingAfter = 0D };
        if (frame == "flow") document.Paragraph(p => p.Runs(runs), style: style);
        else if (frame == "column") document.Content.Row(row => row.RelativeColumn(column => column.Paragraph(p => p.Runs(runs), style: style)));
        else document.Table(new[] { new[] { new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 8D, lineSpacing: style.LineSpacing) }) } },
            style: new PdfTableStyle { FontSize = 8D, HeaderRowCount = 0, CellPaddingX = 0D, CellPaddingY = 0D, BorderColor = null });
        return document.ToBytes();
    }

    private static (double X, double Y, double Width, double Height)[] ReadBands(byte[] bytes, PdfColor? color = null) {
        const string number = @"-?\d+(?:\.\d+)?";
        PdfColor fill = color ?? RuleColor;
        string fillCommand = string.Join(" ", new[] { fill.R, fill.G, fill.B }
            .Select(value => value.ToString("0.###", CultureInfo.InvariantCulture))) + " rg";
        string pattern = Regex.Escape(fillCommand) + @"\s+(?<x>" + number + @")\s+(?<y>" + number +
            @")\s+(?<w>" + number + @")\s+(?<h>" + number + @")\s+re\s+f";
        return Regex.Matches(PdfEncoding.Latin1GetString(bytes), pattern, RegexOptions.CultureInvariant).Cast<Match>()
            .Select(match => (Parse(match, "x"), Parse(match, "y"), Parse(match, "w"), Parse(match, "h"))).ToArray();
    }

    private static double Parse(Match match, string group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
}
