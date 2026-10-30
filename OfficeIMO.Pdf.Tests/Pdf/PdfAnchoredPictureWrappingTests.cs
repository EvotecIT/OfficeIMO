using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfWord = UglyToad.PdfPig.Content.Word;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAnchoredPictureWrappingTests {
    private const string WideWord = "MMMMMMMMM";
    private static readonly string[] Tokens = new[] { WideWord }
        .Concat(Enumerable.Range(1, 24).Select(index => "T" + index.ToString("00")))
        .Concat(new[] { "END" }).ToArray();

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse)]
    [InlineData(PdfTextWhitespaceMode.Preserve)]
    public void SourceLeaderReResolvesParagraphStopWhenItsFontMovesThePictureFrame(PdfTextWhitespaceMode whitespace) {
        var style = TextStyle();
        style.LineSpacing = null;
        style.LineHeight = 1.2D;
        style.TextWhitespaceMode = whitespace;
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(90D), leader: PdfTabLeaderStyle.Dots);
        var content = new PdfPageCanvas();
        content.Image(PdfPngTestImages.CreateRgbPng(2, 2), 0D, 0D, 96D, 48D);
        style.AnchoredCanvas = new PdfCanvasBlock(new PdfCanvasItem[] {
            new PdfCanvasAnchoredImageItem(new object(), new PdfCanvasBlock(content.Items), 96, 48,
                new PdfAnchoredImagePosition(120, 18, false, 30, true),
                new PdfImageTextExclusion(PdfImageTextWrapKind.Square, PdfImageTextWrapSide.BothSides, 0, 0, 0, 0))
        });
        var document = PdfDocument.Create(Options()).Paragraph(p => p.Runs(new[] {
            new PdfTextRun("\t", fontSize: 36D, font: PdfStandardFont.Courier,
                tabLeader: PdfTabLeaderStyle.Dots),
            PdfTextRun.Normal("X", fontSize: 12D)
        }), style: style);
        using var pdf = ReadEvidence("source-leader-picture-frame-" + whitespace, document);
        var page = Assert.Single(pdf.GetPages());
        var following = Assert.Single(page.Letters, letter => letter.Value == "X");
        Assert.Equal(120D, following.StartBaseLine.X, 3);
        var leaders = page.Letters.Where(letter => letter.Value == ".").ToArray();
        Assert.NotEmpty(leaders);
        Assert.All(leaders, letter => {
            Assert.Contains("Courier", letter.FontName, StringComparison.Ordinal);
            Assert.Equal(36D, letter.FontSize, 3);
            Assert.InRange(letter.StartBaseLine.X, 30D, following.StartBaseLine.X);
            Assert.Equal(following.StartBaseLine.Y, letter.StartBaseLine.Y, 3);
        });
        Assert.Equal("X", string.Concat(page.Letters.Where(letter => letter.Value != ".").Select(letter => letter.Value)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SquarePictureUsesBothIntervalsOnOneBaselineAndPaintsEveryTokenOnce(bool behind) {
        var document = PdfDocument.Create(Options()).Paragraph(p => p.Text(string.Join(" ", Tokens)),
            style: PictureStyle(PdfImageTextWrapKind.Square, behind));
        using var pdf = ReadEvidence(behind ? "square-behind" : "square-front", document);
        var page = Assert.Single(pdf.GetPages());
        var image = Assert.Single(page.GetImages());
        Assert.Equal(120D, image.BoundingBox.Left, 3);
        Assert.Equal(370D, image.BoundingBox.Top, 3);
        Assert.Equal(96D, image.BoundingBox.Width, 3);
        Assert.Equal(48D, image.BoundingBox.Height, 3);
        PdfWord[] words = page.GetWords().ToArray();
        AssertTokensOnce(words, Tokens);
        Assert.InRange(Assert.Single(words, word => word.Text == WideWord).BoundingBox.Left, 222D, 330D);
        Assert.Contains(words.GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 3)),
            row => row.Any(word => word.BoundingBox.Left < 114D) &&
                row.Any(word => word.BoundingBox.Left >= 222D));
        Assert.All(words, word => {
            Assert.InRange(word.BoundingBox.Left, 30D, 330D);
            Assert.InRange(word.BoundingBox.Right, 30D, 330D);
            bool alongside = word.BoundingBox.Top > image.BoundingBox.Bottom - 6D &&
                word.BoundingBox.Bottom < image.BoundingBox.Top + 6D;
            if (alongside) Assert.True(word.BoundingBox.Right <= 114.001D || word.BoundingBox.Left >= 221.999D);
        });
    }

    [Fact]
    public void FragmentBarsAndTextTabsKeepParagraphCoordinatesAndTerminalSpacing() {
        PdfDocument Document(bool bar) {
            PdfParagraphStyle style = PictureStyle(PdfImageTextWrapKind.Square);
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6D);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(8D);
            style.AddTabStop(global::OfficeIMO.OfficeLength.Points(240D));
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(36D), PdfTabAlignment.Bar);
            return PdfDocument.Create(Options())
                .Paragraph(p => p.Text("BEFORE"), style: TextStyle())
                .Paragraph(p => p.Text(WideWord + "\nLEFT\tTAB"), style: style)
                .Paragraph(p => p.Text("AFTER"), style: TextStyle());
        }
        using var control = ReadEvidence("fragment-bars-control", Document(false));
        using var actual = ReadEvidence("fragment-bars", Document(true));
        var page = Assert.Single(actual.GetPages());
        var expectedLetters = control.GetPage(1).Letters;
        var actualLetters = page.Letters;
        Assert.Equal(expectedLetters.Count, actualLetters.Count);
        for (int index = 0; index < expectedLetters.Count; index++) {
            Assert.Equal(expectedLetters[index].Value, actualLetters[index].Value);
            Assert.Equal(expectedLetters[index].StartBaseLine.X, actualLetters[index].StartBaseLine.X, 3);
            Assert.Equal(expectedLetters[index].StartBaseLine.Y, actualLetters[index].StartBaseLine.Y, 3);
        }
        var words = page.GetWords().ToArray();
        AssertTokensOnce(words, new[] { "BEFORE", WideWord, "LEFT", "TAB", "AFTER" });
        PdfWord before = Assert.Single(words, word => word.Text == "BEFORE");
        PdfWord first = Assert.Single(words, word => word.Text == WideWord);
        PdfWord left = Assert.Single(words, word => word.Text == "LEFT");
        PdfWord tab = Assert.Single(words, word => word.Text == "TAB");
        PdfWord after = Assert.Single(words, word => word.Text == "AFTER");
        // The first row skips the narrow left interval; its bar keeps the paragraph anchor.
        Assert.Equal(222D, first.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(30D, left.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(270D, tab.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(left.Letters[0].StartBaseLine.Y, tab.Letters[0].StartBaseLine.Y, 3);
        Assert.Equal(21D, before.Letters[0].StartBaseLine.Y - first.Letters[0].StartBaseLine.Y, 3);
        Assert.Equal(15D, first.Letters[0].StartBaseLine.Y - left.Letters[0].StartBaseLine.Y, 3);
        Assert.Equal(23D, left.Letters[0].StartBaseLine.Y - after.Letters[0].StartBaseLine.Y, 3);
        var bars = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && Math.Abs(bounds.Value.Width) < .001D && bounds.Value.Height > 1D)
            .Select(bounds => (X: bounds!.Value.Left, Top: bounds.Value.Top, Height: bounds.Value.Height))
            .OrderByDescending(bounds => bounds.Top).ToArray();
        Assert.Equal(2, bars.Length);
        Assert.All(bars, bar => Assert.Equal(66D, bar.X, 3));
        Assert.Equal(21D, bars[0].Height, 3);
        Assert.Equal(23D, bars[1].Height, 3);
        Assert.Equal(bars[0].Height, bars[0].Top - bars[1].Top, 3);
    }

    [Fact]
    public void FragmentedLinkedIdentifierKeepsOneLogicalTextReplacementAndPositiveLinks() {
        const string source = "ABCD\u200BEFGH\u200BIJKL\u200BMNOP\u200BQRST\u200BUVWX";
        const string expected = "ABCDEFGHIJKLMNOPQRSTUVWX";
        const string uri = "https://example.test/identifier";
        byte[] bytes = PdfDocument.Create(Options()).Paragraph(p => p.Link(source, uri),
            style: PictureStyle(PdfImageTextWrapKind.Square)).ToBytes();
        using var pdf = ReadEvidence("fragment-identifier", bytes);
        Assert.Single(pdf.GetPage(1).GetImages());
        Assert.Equal(expected, PdfReadDocument.Open(bytes).ExtractText().Trim());
        var links = PdfInspector.Inspect(bytes).LinkAnnotations;
        Assert.NotEmpty(links);
        Assert.All(links, link => {
            Assert.Equal(uri, link.Uri);
            Assert.True(link.Width > 0D);
            Assert.True(link.Height > 0D);
            Assert.InRange(link.X1, 30D, 330D);
            Assert.InRange(link.X2, 30D, 330D);
        });
        Assert.Contains(links.GroupBy(link => Math.Round(link.Y1, 3)),
            row => row.Any(link => link.X1 < 114D) && row.Any(link => link.X1 >= 222D));
    }

    [Fact]
    public void TopAndBottomClearanceUsesSourceDistanceOnceAndKeepsPictureAtParagraphOrigin() {
        var wrapped = PdfDocument.Create(Options()).Paragraph(p => p.Text("START T01 T02 T03"),
            style: PictureStyle(PdfImageTextWrapKind.TopAndBottom));
        var control = PdfDocument.Create(Options()).Spacer(48D + 6D)
            .Paragraph(p => p.Text("START T01 T02 T03"), style: TextStyle());
        using var actual = ReadEvidence("top-bottom", wrapped);
        using var expected = ReadEvidence("top-bottom-control", control);
        var page = Assert.Single(actual.GetPages());
        var image = Assert.Single(page.GetImages());
        Assert.Equal(370D, image.BoundingBox.Top, 3);
        foreach (string token in new[] { "START", "T01", "T02", "T03" }) {
            PdfWord word = Assert.Single(page.GetWords(), item => item.Text == token);
            PdfWord reference = Assert.Single(expected.GetPage(1).GetWords(), item => item.Text == token);
            Assert.Equal(reference.Letters[0].StartBaseLine.X, word.Letters[0].StartBaseLine.X, 3);
            Assert.Equal(reference.Letters[0].StartBaseLine.Y, word.Letters[0].StartBaseLine.Y, 3);
        }
    }

    [Fact]
    public void HardBreakAdvancesPhysicalRowInsteadOfFillingTheNextInterval() {
        var document = PdfDocument.Create(Options()).Paragraph(p => p.Text("ALPHA").LineBreak().Text("BETA"),
            style: PictureStyle(PdfImageTextWrapKind.Square));
        using var pdf = ReadEvidence("hard-break", document);
        var words = pdf.GetPage(1).GetWords().ToArray();
        PdfWord alpha = Assert.Single(words, word => word.Text == "ALPHA");
        PdfWord beta = Assert.Single(words, word => word.Text == "BETA");
        Assert.Equal(30D, alpha.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(alpha.Letters[0].StartBaseLine.X, beta.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(15D, alpha.Letters[0].StartBaseLine.Y - beta.Letters[0].StartBaseLine.Y, 3);
        Assert.Single(pdf.GetPage(1).GetImages());
    }

    [Fact]
    public void FollowingParagraphPictureReflowsThePrecedingOccupiedRow() {
        var document = PdfDocument.Create(Options()).Paragraph(p => p.Text(WideWord), style: TextStyle())
            .Paragraph(p => p.Text("FLOW T01 T02 T03"), style: PictureStyle(PdfImageTextWrapKind.Square));
        using var pdf = ReadEvidence("preceding-row", document);
        var page = Assert.Single(pdf.GetPages());
        var image = Assert.Single(page.GetImages());
        Assert.Equal(355D, image.BoundingBox.Top, 3);
        PdfWord before = Assert.Single(page.GetWords(), word => word.Text == WideWord);
        Assert.InRange(before.BoundingBox.Left, 222D, 330D);
        PdfWord flow = Assert.Single(page.GetWords(), word => word.Text == "FLOW");
        Assert.Equal(15D, before.Letters[0].StartBaseLine.Y - flow.Letters[0].StartBaseLine.Y, 3);
    }

    [Fact]
    public void FooterGrowthMovesParagraphPictureUpwardAndPreservesTheFinalBaseline() {
        PdfDocument Footer(bool wrap) {
            var document = PdfDocument.Create(Options());
            var style = PictureStyle(wrap ? PdfImageTextWrapKind.Square : null);
            document.Footer(footer => footer.Content(content =>
                content.Paragraph(p => p.Text(string.Join(" ", Tokens)), style: style), 24D));
            return document.Paragraph(p => p.Text("BODY"), style: TextStyle());
        }
        using var wrapped = ReadEvidence("footer-square", Footer(true));
        using var none = ReadEvidence("footer-none", Footer(false));
        var squarePage = Assert.Single(wrapped.GetPages());
        var nonePage = Assert.Single(none.GetPages());
        var squareImage = Assert.Single(squarePage.GetImages());
        var noneImage = Assert.Single(nonePage.GetImages());
        Assert.True(squareImage.BoundingBox.Top > noneImage.BoundingBox.Top);
        PdfWord squareEnd = Assert.Single(squarePage.GetWords(), word => word.Text == "END");
        PdfWord noneEnd = Assert.Single(nonePage.GetWords(), word => word.Text == "END");
        Assert.Equal(noneEnd.Letters[0].StartBaseLine.Y, squareEnd.Letters[0].StartBaseLine.Y, 3);
        AssertTokensOnce(squarePage.GetWords(), Tokens);
        AssertTokensOnce(nonePage.GetWords(), Tokens);
    }

    [Fact]
    public void PageContinuationReflowsWithoutOldPageFragmentGeometry() {
        string[] tokens = Enumerable.Range(1, 60).Select(index => "T" + index.ToString("00")).ToArray();
        var document = PdfDocument.Create(Options(120D)).Paragraph(p => p.Text(string.Join(" ", tokens)),
            style: PictureStyle(PdfImageTextWrapKind.Square));
        using var pdf = ReadEvidence("page-continuation", document);
        Assert.True(pdf.NumberOfPages > 1);
        AssertTokensOnce(pdf.GetPages().SelectMany(page => page.GetWords()), tokens);
        Assert.Single(pdf.GetPage(1).GetImages());
        Assert.Empty(pdf.GetPage(2).GetImages());
        PdfWord first = pdf.GetPage(2).GetWords().OrderByDescending(word => word.Letters[0].StartBaseLine.Y)
            .ThenBy(word => word.BoundingBox.Left).First();
        Assert.Equal(30D, first.Letters[0].StartBaseLine.X, 3);
        Assert.All(pdf.GetPages().SelectMany(page => page.GetWords()), word => {
            Assert.InRange(word.BoundingBox.Left, 30D, 330D);
            Assert.InRange(word.BoundingBox.Right, 30D, 330D);
        });
    }

    [Fact]
    public void RichStoryDiagnosticsDoNotMeasureItsUnpaintedBackupZones() {
        var report = new PdfConversionReport();
        var options = Options();
        options.ReportDiagnosticsTo(report);
        var document = PdfDocument.Create(options);
        document.Header(header => header.Zones(new string('M', 100), null, null)
            .Content(content => content.Paragraph(p => p.Text("HEADER"), style: TextStyle()), 24D));
        document.Paragraph(p => p.Text("BODY"), style: TextStyle());
        using var pdf = ReadEvidence("active-story-diagnostics", document);
        Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "HEADER");
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.Contains(new string('M', 20)));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "header-footer-content-overflow");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedHeadingKeepsItsOutlineAndLinkBoundsOnThePaintedFragments(bool topAndBottom) {
        const string uri = "https://example.test/heading";
        string title = string.Join(" ", Tokens);
        var options = Options();
        options.CreateOutlineFromHeadings = true;
        var style = new PdfHeadingStyle {
            FontSize = 12D, Bold = false, LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(15D)),
            SpacingBefore = 0D, SpacingAfter = 0D, KeepWithNext = false,
            AnchoredCanvas = PictureStyle(topAndBottom ? PdfImageTextWrapKind.TopAndBottom :
                PdfImageTextWrapKind.Square).AnchoredCanvas
        };
        var document = PdfDocument.Create(options);
        document.Content.Heading(1, title, linkUri: uri, style: style);
        byte[] bytes = document.ToBytes();
        using var pdf = ReadEvidence(topAndBottom ? "heading-top-bottom" : "heading-square", bytes);
        var page = Assert.Single(pdf.GetPages());
        var image = Assert.Single(page.GetImages()).BoundingBox;
        AssertTokensOnce(page.GetWords(), Tokens);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfOutlineItem outline = Assert.Single(info.Outlines);
        Assert.Equal(title, outline.Title);
        Assert.Equal(1, outline.Level);
        Assert.Equal(1, outline.PageNumber);
        Assert.Equal(topAndBottom ? image.Bottom - 6D : 370D, outline.DestinationTop!.Value, 3);
        Assert.NotEmpty(info.LinkAnnotations);
        Assert.All(info.LinkAnnotations, link => {
            Assert.Equal(uri, link.Uri);
            Assert.True(link.Width > 0D && link.Height > 0D);
            Assert.InRange(link.X1, 30D, 330D);
            Assert.InRange(link.X2, 30D, 330D);
            bool alongside = link.Y2 > image.Bottom - 6D && link.Y1 < image.Top + 6D;
            Assert.True(!alongside || link.X2 <= image.Left - 6D + .01D ||
                link.X1 >= image.Right + 6D - .01D);
            if (topAndBottom) Assert.True(link.Y2 <= image.Bottom - 6D + .01D);
        });
        if (!topAndBottom) Assert.Contains(info.LinkAnnotations.GroupBy(link => Math.Round(link.Y1, 3)),
            row => row.Any(link => link.X1 < image.Left - 6D) && row.Any(link => link.X1 >= image.Right + 6D));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void NonConvergingRunningContentReportsOnlyUnstablePicturePlacement(bool withPicture, bool movingPicture) {
        const double headerDistance = 18D;
        double tallHeader = Options().PageHeight.ToPoints() - Options().MarginBottom.ToPoints() - headerDistance - 2D * 15D;
        PdfDocument Document(Func<PdfRunningContentContext, double> headerHeight) => PdfDocument.Create(pdf => {
            pdf.Section(page => page.Content(content => content.Paragraph(p => p.Text("ANCHOR"),
                style: withPicture && !movingPicture
                    ? PictureStyle(PdfImageTextWrapKind.Square, paragraphY: false) : TextStyle())));
            pdf.Section(page => {
                page.Header(header => header.Content(context =>
                    content => content.Spacer(headerHeight(context)), headerDistance));
                page.Content(content => {
                    for (int row = 1; row <= 7; row++)
                        content.Paragraph(p => p.Text("ROW" + row), style: withPicture && movingPicture && row == 1
                            ? PictureStyle(PdfImageTextWrapKind.Square) : TextStyle());
                });
            });
        }, Options());

        string profile = !withPicture ? "none" : movingPicture ? "moving" : "stable";
        using var small = ReadEvidence("nonconvergence-small-" + profile, Document(_ => 0D));
        using var large = ReadEvidence("nonconvergence-large-" + profile, Document(_ => tallHeader));
        Assert.Equal(2, small.NumberOfPages);
        Assert.True(large.NumberOfPages > small.NumberOfPages);
        string[] expected = new[] { "ANCHOR" }.Concat(Enumerable.Range(1, 7).Select(row => "ROW" + row)).ToArray();
        AssertTokensOnce(small.GetPages().SelectMany(page => page.GetWords()), expected);
        AssertTokensOnce(large.GetPages().SelectMany(page => page.GetWords()), expected);
        if (withPicture) {
            var first = Assert.Single(small.GetPages().SelectMany(page => page.GetImages())).BoundingBox;
            var second = Assert.Single(large.GetPages().SelectMany(page => page.GetImages())).BoundingBox;
            Assert.Equal(first.Left, second.Left, 3);
            Assert.Equal(first.Width, second.Width, 3);
            Assert.Equal(first.Height, second.Height, 3);
            if (movingPicture) Assert.NotEqual(Math.Round(first.Top, 3), Math.Round(second.Top, 3));
            else Assert.Equal(first.Top, second.Top, 3);
        } else {
            Assert.Empty(small.GetPages().SelectMany(page => page.GetImages()));
            Assert.Empty(large.GetPages().SelectMany(page => page.GetImages()));
        }
        // A pure function of the advertised document page count alternates the two
        // proven endpoint layouts; no call counter, randomness or timing is used.
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => Document(context =>
            context.DocumentPages <= small.NumberOfPages ? tallHeader : 0D).ToBytes());
        Assert.StartsWith(movingPicture ? "Anchored picture layout" : "Running PDF content", error.Message);
    }

    private static PdfCreateOptions Options(double height = 400D) => new PdfCreateOptions {
        PageWidth = global::OfficeIMO.OfficeLength.Points(360D), PageHeight = global::OfficeIMO.OfficeLength.Points(height), MarginLeft = global::OfficeIMO.OfficeLength.Points(30D), MarginRight = global::OfficeIMO.OfficeLength.Points(30D),
        MarginTop = global::OfficeIMO.OfficeLength.Points(30D), MarginBottom = global::OfficeIMO.OfficeLength.Points(30D), DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D), ShowHeader = false, ShowPageNumbers = false
    };

    private static PdfParagraphStyle TextStyle() => new PdfParagraphStyle {
        FontSize = global::OfficeIMO.OfficeLength.Points(12D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(15D)),
        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), WidowControl = false
    };

    private static PdfParagraphStyle PictureStyle(PdfImageTextWrapKind? kind, bool behind = false, bool paragraphY = true) {
        PdfParagraphStyle style = TextStyle();
        var content = new PdfPageCanvas();
        if (behind) content.Image(PdfPngTestImages.CreateRgbPng(2, 2), 0D, 0D, 96D, 48D);
        else content.ForegroundImage(PdfPngTestImages.CreateRgbPng(2, 2), 0D, 0D, 96D, 48D);
        IReadOnlyList<PdfCanvasItem> paint = behind
            ? new PdfCanvasItem[] { new PdfCanvasBehindTextItem(content.Items, 0L) } : content.Items;
        var exclusion = kind.HasValue ? new PdfImageTextExclusion(kind.Value,
            PdfImageTextWrapSide.BothSides, 6D, 6D, 6D, 6D) : null;
        var picture = new PdfCanvasAnchoredImageItem(new object(), new PdfCanvasBlock(paint),
            96D, 48D, new PdfAnchoredImagePosition(120D, paragraphY ? 0D : 30D, false, 30D, paragraphY), exclusion);
        style.AnchoredCanvas = new PdfCanvasBlock(new PdfCanvasItem[] { picture });
        return style;
    }

    private static void AssertTokensOnce(IEnumerable<PdfWord> words, IReadOnlyList<string> expected) {
        PdfWord[] actual = words.ToArray();
        foreach (string token in expected) Assert.Single(actual, word => word.Text == token);
        Assert.Equal(expected.Count, actual.Count(word => expected.Contains(word.Text)));
    }

    private static PdfPigDocument ReadEvidence(string name, PdfDocument document) => ReadEvidence(name, document.ToBytes());

    private static PdfPigDocument ReadEvidence(string name, byte[] bytes) {
        string output = Path.Combine(AppContext.BaseDirectory, "picture-wrap-artifacts");
        Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, name + ".pdf"), bytes);
        return PdfPigDocument.Open(bytes);
    }
}
