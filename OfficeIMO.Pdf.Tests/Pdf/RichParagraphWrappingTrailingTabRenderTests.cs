using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class RichParagraphWrappingTests {
    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, false, true)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true, true)]
    public void ParagraphTabs_TrailingUnderlineLeaderRendersWithoutFollowingText(
        PdfTextWhitespaceMode whitespace, bool hardBreak, bool trailingSpace) {
        const double margin = 36D, firstStop = 72D, leaderStop = 216D;
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) };
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(firstStop));
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(leaderStop), leader: PdfTabLeaderStyle.Underscores);
        byte[] bytes = PdfDocument.Create(TrailingTabOptions())
            .Paragraph(p => {
                p.Tab().Tab();
                if (trailingSpace) p.Text(" ");
                if (hardBreak) p.LineBreak().Text("AFTER");
            }, style: style)
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var leaders = page.Letters.Where(letter => letter.Value == "_")
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.True(leaders.Length >= 3, "A trailing tab must paint its authored underline leader without a following text run.");
        Assert.Equal(margin + firstStop, leaders[0].StartBaseLine.X, 2);
        double glyphAdvance = leaders[0].EndBaseLine.X - leaders[0].StartBaseLine.X;
        Assert.True(glyphAdvance > 0D);
        Assert.InRange(leaders[leaders.Length - 1].EndBaseLine.X,
            margin + leaderStop - glyphAdvance - .02D, margin + leaderStop + .02D);
        Assert.All(leaders, letter => Assert.Equal(leaders[0].StartBaseLine.Y, letter.StartBaseLine.Y, 2));
        var words = page.GetWords().ToArray();
        var next = Assert.Single(words, word => word.Text == "NEXT");
        Assert.Equal(margin, next.Letters[0].StartBaseLine.X, 2);
        Assert.True(next.Letters[0].StartBaseLine.Y < leaders[0].StartBaseLine.Y);
        if (hardBreak) {
            var after = Assert.Single(words, word => word.Text == "AFTER");
            Assert.Equal(margin, after.Letters[0].StartBaseLine.X, 2);
            Assert.True(after.Letters[0].StartBaseLine.Y < leaders[0].StartBaseLine.Y);
            Assert.True(after.Letters[0].StartBaseLine.Y > next.Letters[0].StartBaseLine.Y);
        }
    }

    [Fact]
    public void ParagraphTabs_PlainTrailingTabsDoNotPaintLeadersOrIndentFollowingContent() {
        byte[] bytes = PdfDocument.Create(TrailingTabOptions())
            .Paragraph(p => p.Text("BEFORE").Tab().Text(" "), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .Paragraph(p => p.Tab().Text(" ").LineBreak().Text("AFTER").Tab().Text(" "),
                style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var visible = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("BEFOREAFTERNEXT", string.Concat(visible.Select(letter => letter.Value)));
        var words = page.GetWords().ToArray();
        var before = Assert.Single(words, word => word.Text == "BEFORE");
        var after = Assert.Single(words, word => word.Text == "AFTER");
        var next = Assert.Single(words, word => word.Text == "NEXT");
        foreach (var word in new[] { before, after, next })
            Assert.Equal(36D, word.Letters[0].StartBaseLine.X, 2);
        Assert.True(before.Letters[0].StartBaseLine.Y > after.Letters[0].StartBaseLine.Y);
        Assert.True(after.Letters[0].StartBaseLine.Y > next.Letters[0].StartBaseLine.Y);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_SpaceBeforeNextTabRetainsBothLeaders(
        PdfTextWhitespaceMode whitespace, bool followingText) {
        const double margin = 36D, firstStop = 72D, secondStop = 216D;
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) };
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(firstStop), leader: PdfTabLeaderStyle.Dots);
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(secondStop), leader: PdfTabLeaderStyle.Underscores);
        PdfCreateOptions options = TrailingTabOptions();
        options.DefaultFont = PdfStandardFont.Courier;
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => {
            p.Text("A").Tab().Text(" ").Tab();
            if (followingText) p.Text("B");
        }, style: style).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var a = Assert.Single(page.Letters, letter => letter.Value == "A");
        var dots = page.Letters.Where(letter => letter.Value == ".").ToArray();
        var underscores = page.Letters.Where(letter => letter.Value == "_")
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.NotEmpty(dots);
        Assert.NotEmpty(underscores);
        Assert.Equal(margin, a.StartBaseLine.X, 2);
        double monospaceAdvance = a.EndBaseLine.X - a.StartBaseLine.X;
        Assert.True(monospaceAdvance > 0D);
        double expectedStart = margin + firstStop +
            (whitespace == PdfTextWhitespaceMode.Collapse ? 0D : monospaceAdvance);
        Assert.Equal(expectedStart, underscores[0].StartBaseLine.X, 2);
        Assert.All(dots.Concat(underscores), letter => Assert.Equal(a.StartBaseLine.Y, letter.StartBaseLine.Y, 2));
        Assert.Equal(followingText ? "AB" : "A", string.Concat(page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value) && letter.Value != "." && letter.Value != "_")
            .Select(letter => letter.Value)));
        if (followingText) {
            var b = Assert.Single(page.Letters, letter => letter.Value == "B");
            Assert.Equal(margin + secondStop, b.StartBaseLine.X, 2);
            Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 2);
        }
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_NoLeaderAdvancesTextAndInlineContentWithoutPaintingSpace(
        PdfTextWhitespaceMode whitespace, bool inline) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) };
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(180D));
        PdfTextRun content = inline
            ? PdfTextRun.Inline(new PdfInlineBox(10D, 10D, background: global::OfficeIMO.OfficeColor.Black))
            : PdfTextRun.Normal("VISIBLE", fontSize: global::OfficeIMO.OfficeLength.Points(20D));
        byte[] bytes = PdfDocument.Create(TrailingTabInkOptions())
            .Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D)), content }), style: style)
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        // SPACE has an outline in this font. Retain every extracted glyph so
        // a manufactured separator cannot disappear behind a whitespace filter.
        Assert.Equal(inline ? "NEXT" : "VISIBLENEXT", string.Concat(page.Letters.Select(letter => letter.Value)));
        if (inline) {
            var bounds = Assert.Single(page.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), rectangle => rectangle.HasValue);
            Assert.Equal(216D, bounds!.Value.Left, 3);
            Assert.Equal(10D, bounds.Value.Width, 3);
            Assert.Equal(10D, bounds.Value.Height, 3);
        } else {
            var visible = Assert.Single(page.Letters, letter => letter.Value == "V");
            Assert.Equal(216D, visible.StartBaseLine.X, 3);
            Assert.Equal(20D, visible.FontSize, 3);
        }
        var next = Assert.Single(page.Letters, letter => letter.Value == "N");
        Assert.Equal(36D, next.StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_TerminalLeaderUsesItsPaintedFontBox(
        PdfTextWhitespaceMode whitespace, bool trailingSpace) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) };
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(180D), leader: PdfTabLeaderStyle.Underscores);
        byte[] leaderBytes = PdfDocument.Create(TrailingTabInkOptions())
            .Paragraph(p => {
                // The effective paragraph stop overrides the run's dotted leader.
                p.Runs(new[] { new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D), tabLeader: PdfTabLeaderStyle.Dots) });
                if (trailingSpace) p.Text(" ");
            }, style: style)
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .ToBytes();
        byte[] referenceBytes = PdfDocument.Create(TrailingTabInkOptions())
            .Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("___", fontSize: global::OfficeIMO.OfficeLength.Points(24D)) }),
                style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
            .ToBytes();

        using var leaderPdf = UglyToad.PdfPig.PdfDocument.Open(leaderBytes);
        using var referencePdf = UglyToad.PdfPig.PdfDocument.Open(referenceBytes);
        var leaderPage = Assert.Single(leaderPdf.GetPages());
        var referencePage = Assert.Single(referencePdf.GetPages());
        var leaders = leaderPage.Letters.Where(letter => letter.Value == "_").ToArray();
        Assert.True(leaders.Length >= 3);
        Assert.Equal(new string('_', leaders.Length) + "NEXT", string.Concat(leaderPage.Letters.Select(letter => letter.Value)));
        var referenceLeader = referencePage.Letters.First(letter => letter.Value == "_");
        Assert.Equal(36D, leaders[0].StartBaseLine.X, 3);
        Assert.All(leaders, letter => {
            Assert.Equal(24D, letter.FontSize, 3);
            Assert.Equal(referenceLeader.StartBaseLine.Y, letter.StartBaseLine.Y, 3);
        });
        var next = Assert.Single(leaderPage.Letters, letter => letter.Value == "N");
        var referenceNext = Assert.Single(referencePage.Letters, letter => letter.Value == "N");
        Assert.Equal(referenceNext.StartBaseLine.Y, next.StartBaseLine.Y, 3);
        Assert.Equal(36D, next.StartBaseLine.X, 3);
    }

    [Fact]
    public void ParagraphTabs_SubThresholdLeaderKeepsThePlainTabFontBox() {
        byte[] Render(PdfTabLeaderStyle leader) {
            var style = new PdfParagraphStyle { TextWhitespaceMode = PdfTextWhitespaceMode.Preserve, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) };
            // At24pt this12pt gap is smaller than the existing three-glyph
            // leader threshold. The selected leader therefore paints no glyphs.
            style.AddTabStop(global::OfficeIMO.OfficeLength.Points(12D), leader: leader);
            return PdfDocument.Create(TrailingTabInkOptions())
                .Paragraph(p => p.Runs(new[] { new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D), tabLeader: leader) }), style: style)
                .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
                .ToBytes();
        }
        using var leaderPdf = UglyToad.PdfPig.PdfDocument.Open(Render(PdfTabLeaderStyle.Underscores));
        using var plainPdf = UglyToad.PdfPig.PdfDocument.Open(Render(PdfTabLeaderStyle.None));
        var leaderPage = Assert.Single(leaderPdf.GetPages());
        var plainPage = Assert.Single(plainPdf.GetPages());
        Assert.Equal("NEXT", string.Concat(leaderPage.Letters.Select(letter => letter.Value)));
        Assert.Equal("NEXT", string.Concat(plainPage.Letters.Select(letter => letter.Value)));
        for (int index = 0; index < plainPage.Letters.Count; index++) {
            Assert.Equal(plainPage.Letters[index].StartBaseLine.X, leaderPage.Letters[index].StartBaseLine.X, 3);
            Assert.Equal(plainPage.Letters[index].StartBaseLine.Y, leaderPage.Letters[index].StartBaseLine.Y, 3);
        }
    }
    [Fact]
    public void ParagraphTabs_BufferedLeaderRefitsItsSourceFontBeforeSelectingTheNextStop() {
        var requests = new System.Collections.Generic.List<(double Height, double MinimumWidth)>();
        Func<int, double, double, double, (double Width, double Origin, double Gap)> frame =
            (line, completed, height, minimumWidth) => {
                requests.Add((height, minimumWidth));
                return (height > 20D ? 60D : 288D, 0D, 0D);
            };
        var runs = new[] {
            new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D), tabLeader: PdfTabLeaderStyle.Underscores),
            PdfTextRun.Normal(" ", fontSize: global::OfficeIMO.OfficeLength.Points(12D)),
            PdfTextRun.Tab()
        };
        object result = InvokePrivateFontMethod<object>("WrapRichRunsCoreWithFirstLineOrigin",
            runs, 288D, 12D, PdfStandardFont.Helvetica, 14.4D, null!, null!, 36D,
            TrailingTabInkOptions(), new[] {
                new PdfTabStop(global::OfficeIMO.OfficeLength.Points(72D), leader: PdfTabLeaderStyle.Underscores),
                new PdfTabStop(global::OfficeIMO.OfficeLength.Points(180D))
            }, frame, null!, null!, PdfTextWhitespaceMode.Preserve);

        // The existing private-owner harness admits a deterministic exclusion
        // frame: only the actual24pt source box narrows its available width.
        Assert.Contains(requests, request => request.Height >= 28.8D - .001D);
        var line = Assert.Single(ExtractLines(result));
        var leader = Assert.Single(line, segment => ExtractLeadingTabLeader(segment) == PdfTabLeaderStyle.Underscores);
        Assert.Equal(60D, ExtractLeadingAdvance(leader), 3);
        Assert.Equal(24D, (double)leader.GetType().GetProperty("FontSize")!.GetValue(leader)!, 3);
        var heights = (System.Collections.Generic.List<double>)result.GetType().GetField("Item2")!.GetValue(result)!;
        Assert.Equal(28.8D, Assert.Single(heights), 3);
    }
    [Fact]
    public void ParagraphTabs_FinalLeaderFontBoxRemainsAdmittedAfterFloatRefit() {
        var requests = new System.Collections.Generic.List<(double Height, double Origin, double Gap)>();
        Func<int, double, double, double, (double Width, double Origin, double Gap)> frame =
            (line, completed, height, minimumWidth) => {
                // A left float occupies x0..100, starting20pt below this line
                // and ending60pt below it. This matches the production exclusion
                // rule: use the widest interval, or move below if it cannot fit.
                var result = height <= 20D
                    ? (Width: 288D, Origin: 0D, Gap: 0D)
                    : minimumWidth <= 188D
                        ? (Width: 188D, Origin: 100D, Gap: 0D)
                        : (Width: 288D, Origin: 0D, Gap: 60D);
                requests.Add((height, result.Origin, result.Gap));
                return result;
            };
        object result = InvokePrivateFontMethod<object>("WrapRichRunsCoreWithFirstLineOrigin",
            new[] { new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D), tabLeader: PdfTabLeaderStyle.Underscores) },
            288D, 12D, PdfStandardFont.Helvetica, 14.4D, null!, null!, 36D,
            TrailingTabInkOptions(), new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(72D), leader: PdfTabLeaderStyle.Underscores) },
            frame, null!, null!, PdfTextWhitespaceMode.Preserve);
        var line = Assert.Single(ExtractLines(result));
        var tab = Assert.Single(line);
        double gap = ExtractLeadingAdvance(tab);
        Assert.Equal(72D, gap, 3);
        Assert.Equal(PdfTabLeaderStyle.Underscores, ExtractLeadingTabLeader(tab));
        Assert.Equal(24D, (double)tab.GetType().GetProperty("FontSize")!.GetValue(tab)!, 3);
        var heights = (System.Collections.Generic.List<double>)result.GetType().GetField("Item2")!.GetValue(result)!;
        var admitted = requests.Last();
        // A final paintable24pt leader cannot use a callback box that was only
        // admitted for the12pt paragraph. Clearance is separate from its box.
        Assert.Equal(admitted.Height + admitted.Gap, Assert.Single(heights), 3);
        if (gap >= 36D) Assert.Equal(28.8D, admitted.Height, 3);
    }

    [Theory]
    [InlineData(false, 72D)]
    [InlineData(true, 72D)]
    [InlineData(false, 112D)]
    [InlineData(true, 112D)]
    public void ParagraphTabs_WidowCarriedLeaderRefreshesItsFontBox(bool exact, double stop) {
        byte[] Render(int count, bool floating) {
            PdfCreateOptions options = TrailingTabInkOptions();
            options.PageWidth = global::OfficeIMO.OfficeLength.Points(400D); options.PageHeight = global::OfficeIMO.OfficeLength.Points(240D);
            options.MarginLeft = options.MarginRight = options.MarginTop = options.MarginBottom = global::OfficeIMO.OfficeLength.Points(40D);
            var document = PdfDocument.Create(options);
            if (floating) document.Table(new[] { new[] { "T" } }, style: new PdfTableStyle {
                HeaderRowCount = 0, FontSize = global::OfficeIMO.OfficeLength.Points(2D), CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D),
                ColumnWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100D) },
                MinRowHeight = global::OfficeIMO.OfficeLength.Points(145D), BorderColor = null, Position = new PdfTablePosition()
            });
            var style = new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12D), LineHeight = 1.2D, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
                TextWhitespaceMode = PdfTextWhitespaceMode.Preserve, WidowControl = true,
                LineSpacing = exact ? PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(14.4D)) : null
            };
            style.AddTabStop(global::OfficeIMO.OfficeLength.Points(stop), leader: PdfTabLeaderStyle.Underscores);
            var runs = new System.Collections.Generic.List<PdfTextRun>();
            for (int index = 0; index < count; index++) {
                runs.Add(new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D)));
                if (index + 1 < count) runs.Add(PdfTextRun.Normal("\n", fontSize: global::OfficeIMO.OfficeLength.Points(12D)));
            }
            return document.Paragraph(p => p.Runs(runs), style: style)
                .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D) })
                .ToBytes();
        }
        using var carried = UglyToad.PdfPig.PdfDocument.Open(Render(12, floating: true));
        using var reference = UglyToad.PdfPig.PdfDocument.Open(Render(2, floating: false));
        Assert.Equal(2, carried.NumberOfPages);
        var actualPage = carried.GetPage(2);
        var expectedPage = Assert.Single(reference.GetPages());
        // The left float suppresses the stop on page1. Widow control carries
        // two lines to a clear page, where their source24pt leaders paint.
        // Reanchoring must refresh baseline, occupied box and next-line advance;
        // exact spacing deliberately keeps the authored14.4pt contract.
        Assert.Contains(actualPage.Letters, letter => letter.Value == "_");
        Assert.Equal(string.Concat(expectedPage.Letters.Select(letter => letter.Value)),
            string.Concat(actualPage.Letters.Select(letter => letter.Value)));
        Assert.Equal(expectedPage.Letters.Count, actualPage.Letters.Count);
        for (int index = 0; index < expectedPage.Letters.Count; index++) {
            Assert.Equal(expectedPage.Letters[index].FontSize, actualPage.Letters[index].FontSize, 3);
            Assert.Equal(expectedPage.Letters[index].StartBaseLine.X, actualPage.Letters[index].StartBaseLine.X, 3);
            Assert.Equal(expectedPage.Letters[index].StartBaseLine.Y, actualPage.Letters[index].StartBaseLine.Y, 3);
        }
    }
    [Fact]
    public void ParagraphTabs_ContinuationRetainsTheEffectiveSourceStyle() {
        const string family = "Tab Continuation";
        PdfCreateOptions options = TrailingTabInkOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family,
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 200, 200,
                Enumerable.Range(32, 95).ToArray())));
        PdfTextRun source = new PdfTextRun("\t", fontSize: global::OfficeIMO.OfficeLength.Points(24D), font: PdfStandardFont.Helvetica,
            fontFamily: family, baseline: PdfTextBaseline.Superscript,
            tabLeader: PdfTabLeaderStyle.Underscores, tabAlignment: PdfTabAlignment.Right)
            .WithFeatureSettings(OfficeIMO.Drawing.OfficeTextFeatureSettings.Default.With("kern", 0))
            .WithHorizontalTextScaling(90D).WithCharacterSpacing(global::OfficeIMO.OfficeLength.Points(.3D));
        object wrapped = InvokeWrapRichRunsWithOptions(new[] { source }, 288D, 12D, PdfStandardFont.Helvetica, options);
        object lines = wrapped.GetType().GetField("Item1")!.GetValue(wrapped)!;
        var restored = InvokePrivateFontMethod<System.Collections.Generic.List<PdfTextRun>>(
            "BuildTextRunsFromWrappedLines", lines, 0, 1);
        PdfTextRun tab = Assert.Single(restored, run => run.Text == "\t");
        Assert.Equal((source.FontSize?.ToPoints()), (tab.FontSize?.ToPoints()));
        Assert.Equal(source.Font, tab.Font);
        Assert.Equal(source.FontFamily, tab.FontFamily);
        Assert.Equal(source.Baseline, tab.Baseline);
        Assert.Equal(source.TabLeader, tab.TabLeader);
        Assert.Equal(source.TabAlignment, tab.TabAlignment);
        Assert.Equal(source.FeatureSettings, tab.FeatureSettings);
        Assert.Equal(source.HorizontalTextScaling, tab.HorizontalTextScaling);
        Assert.Equal((source.CharacterSpacing.ToPoints()), (tab.CharacterSpacing.ToPoints()));
    }
    [Fact]
    public void ParagraphTabs_RelativeContinuationRetainsItsFrameBinding() {
        PdfTextRun source = PdfTextRun.RelativeTab(PdfTabAlignment.Right,
            PdfTabLeaderStyle.Underscores, .5D, -4D);
        PdfCreateOptions options = TrailingTabInkOptions();
        object wrapped = InvokeWrapRichRunsWithOptions(new[] { source }, 288D, 24D, PdfStandardFont.Helvetica, options);
        object lines = wrapped.GetType().GetField("Item1")!.GetValue(wrapped)!;
        var restored = InvokePrivateFontMethod<System.Collections.Generic.List<PdfTextRun>>(
            "BuildTextRunsFromWrappedLines", lines, 0, 1);
        PdfTextRun tab = Assert.Single(restored, run => run.Text == "\t");
        Assert.Equal(source.RelativeTabStop, tab.RelativeTabStop);
        Assert.Equal(24D, (tab.FontSize?.ToPoints()));
        object resized = InvokeWrapRichRunsWithOptions(restored, 200D, 12D, PdfStandardFont.Helvetica, options);
        var resizedTab = Assert.Single(Assert.Single(ExtractLines(resized)));
        Assert.Equal(96D, ExtractLeadingAdvance(resizedTab), 3);
    }
    private static PdfCreateOptions TrailingTabInkOptions() {
        PdfCreateOptions options = TrailingTabOptions();
        options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 200, 200,
                Enumerable.Range(32, 95).ToArray())));
        return options;
    }
    private static PdfCreateOptions TrailingTabOptions() => new PdfCreateOptions {
        BackgroundColor = null,
        PageWidth = global::OfficeIMO.OfficeLength.Points(360D), PageHeight = global::OfficeIMO.OfficeLength.Points(220D),
        MarginLeft = global::OfficeIMO.OfficeLength.Points(36D), MarginRight = global::OfficeIMO.OfficeLength.Points(36D), MarginTop = global::OfficeIMO.OfficeLength.Points(36D), MarginBottom = global::OfficeIMO.OfficeLength.Points(36D),
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D)
    };
}
