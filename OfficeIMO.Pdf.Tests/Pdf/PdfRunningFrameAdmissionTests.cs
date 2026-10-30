using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunningFrameAdmissionTests {
    private const double PageHeight = 320D;
    private const double PhysicalMargin = 32D;
    private const double HeaderDistance = 18D;
    private const double SmallBodyHeight = 50D;
    private const double LargeBodyHeight = 200D;
    private const double RequiredParagraphHeight = SmallBodyHeight * 1.5D;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphMovesToLargerRunningFrameWithoutAnExplicitBreak(bool keepTogether) {
        PdfDocument document = CreateDocument(keepTogether, explicitBreak: false);

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether, targetPage: 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitBreakProvesTheParagraphFitsTheLaterRunningFrame(bool keepTogether) {
        PdfDocument document = CreateDocument(keepTogether, explicitBreak: true);

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether, targetPage: 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EqualLargeRunningFramesKeepTheBodyOnOnePage(bool keepTogether) {
        PdfDocument document = CreateDocument(keepTogether, explicitBreak: false,
            firstBodyHeight: LargeBodyHeight);

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether, targetPage: 1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphRejectsFramesThatCannotFitInsteadOfExhaustingThePageLimit(bool keepTogether) {
        PdfDocument document = CreateDocument(keepTogether, explicitBreak: false,
            laterBodyHeight: SmallBodyHeight);

        // The page limit bounds a faulty retry path. A fit failure must remain
        // ArgumentException rather than running into the generated-page limit.
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphReachesTheLargerOddFrameAfterShortFirstAndEvenFrames(bool keepTogether) {
        PdfDocument document = CreateDocument(keepTogether, explicitBreak: false,
            evenBodyHeight: SmallBodyHeight);

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether, targetPage: 3);
        Assert.Empty(pdf.GetPage(2).Text);
    }

    [Fact]
    public void PaddedKeptContainerReachesTheLargerRunningFrame() {
        double padding = SmallBodyHeight / 10D;
        PdfDocument document = CreateContentDocument(content => content.Panel(panel =>
            panel.Paragraph(paragraph => paragraph.Text("PANEL"), style: TargetStyle(keepTogether: false)),
            new PdfPanelStyle {
                PaddingX = padding,
                PaddingY = padding,
                SpacingBefore = 0D,
                SpacingAfter = 0D,
                KeepTogether = true
            }));

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Equal("PANEL", pdf.GetPage(2).Text);
    }

    [Fact]
    public void StrictKeptTableReachesTheLargerRunningFrame() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = style.CellPaddingY = 0D;
        style.SpacingBefore = style.SpacingAfter = 0D;
        style.FixedRowHeights = new List<double?> { RequiredParagraphHeight };
        style.KeepTogether = true;
        style.AllowRowBreakAcrossPages = false;
        PdfDocument document = CreateContentDocument(content =>
            content.Table(new[] { new[] { "TABLE" } }, style: style));

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Equal("TABLE", pdf.GetPage(2).Text);
    }

    [Fact]
    public void ContextualRunningFramesReachALaterFitWithoutEagerCallbacks() {
        var headerPages = new List<int>();
        bool lazyBodyMaterialized = false;
        PdfDocument document = CreateContentDocument(content => content.Deferred(context => {
            Assert.NotEmpty(headerPages);
            Assert.Equal(1, headerPages[headerPages.Count - 1]);
            lazyBodyMaterialized = true;
            return body => body.Paragraph(paragraph => paragraph.Text("TARGET"),
                style: TargetStyle(keepTogether: false));
        }), configureHeader: header => header.Content(context => {
            if (context.PageNumber > 1) Assert.True(lazyBodyMaterialized);
            headerPages.Add(context.PageNumber);
            double bodyHeight = context.PageNumber < 3 ? SmallBodyHeight : LargeBodyHeight;
            return story => story.Spacer(HeaderHeight(bodyHeight));
        }, HeaderDistance));

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether: false, targetPage: 3);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.True(lazyBodyMaterialized);
        // Stabilization may replay the document. Every pass must visit actual
        // physical pages in order, with one callback per page and no lookahead.
        Assert.NotEmpty(headerPages);
        Assert.Equal(0, headerPages.Count % 3);
        for (int index = 0; index < headerPages.Count; index += 3)
            Assert.Equal(new[] { 1, 2, 3 }, headerPages.Skip(index).Take(3));
    }

    [Fact]
    public void UnboundedContextualFramesRetainConservativeFitRejection() {
        using var cancellation = new CancellationTokenSource();
        PdfDocument document = CreateContentDocument(content =>
            content.Paragraph(paragraph => paragraph.Text("TARGET"), style: TargetStyle(keepTogether: false)),
            configureHeader: header => header.Content(context => {
                // A deterministic watchdog makes an erroneous unlimited search
                // fail by cancellation without a host-dependent time budget.
                if (context.PageNumber > 2) cancellation.Cancel();
                return story => story.Spacer(HeaderHeight(SmallBodyHeight));
            }, HeaderDistance), maximumPages: null);

        Assert.Throws<ArgumentException>(() => document.ToBytes(cancellationToken: cancellation.Token));
    }

    [Fact]
    public void HeadingLineReachesTheLargerOddFrameAfterShortFirstAndEvenFrames() {
        PdfDocument document = CreateContentDocument(content =>
            content.Heading(1, "HEADING", style: new PdfHeadingStyle {
                FontSize = 10D,
                LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(RequiredParagraphHeight)),
                SpacingBefore = 0D,
                SpacingAfter = 0D,
                KeepWithNext = false
            }), header => ConfigureStaticHeader(header, SmallBodyHeight, LargeBodyHeight, SmallBodyHeight));

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Equal("HEADING", pdf.GetPage(3).Text);
    }

    [Fact]
    public void KeptListReachesTheLargerRunningFrame() {
        PdfDocument document = CreateContentDocument(content =>
            content.Bullets(new[] { "LIST" }, style: new PdfListStyle {
                FontSize = 10D,
                LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(RequiredParagraphHeight)),
                SpacingBefore = 0D,
                SpacingAfter = 0D,
                ItemSpacing = 0D,
                KeepTogether = true,
                KeepTogetherOverflow = PdfKeepTogetherOverflowBehavior.Reject
            }));

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Single(Regex.Matches(pdf.GetPage(2).Text, "LIST").Cast<Match>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KeptFlowReachesTheLargerRunningFrame(bool replayable) {
        Action<PdfContentBuilder> target = flow => flow.Paragraph(paragraph => paragraph.Text("FLOW"),
            style: TargetStyle(keepTogether: false));
        var options = new PdfFlowOptions { KeepTogether = true };
        PdfDocument document = CreateContentDocument(content => {
            if (replayable) content.Deferred(context => target, options);
            else content.Flow(target, options);
        });

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Equal("FLOW", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void KeptFlowRechecksRunningFramesAfterAnOrdinaryMove(bool replayable, bool laterFrameFits) {
        Action<PdfContentBuilder> target = flow => flow.Paragraph(paragraph => paragraph.Text("FLOW"),
            style: TargetStyle(keepTogether: false));
        PdfDocument document = CreateContentDocument(content => {
            content.Spacer(150D);
            var options = new PdfFlowOptions { KeepTogether = true };
            if (replayable) content.Deferred(_ => target, options);
            else content.Flow(target, options);
        }, header => ConfigureStaticHeader(header, LargeBodyHeight,
            laterFrameFits ? LargeBodyHeight : SmallBodyHeight, SmallBodyHeight));

        if (!laterFrameFits) {
            Assert.Throws<ArgumentException>(() => document.ToBytes());
            return;
        }
        AssertLaterRunningFrame(document, "FLOW");
    }

    [Theory]
    [InlineData("kept", true)]
    [InlineData("fixed", true)]
    [InlineData("fixed", false)]
    [InlineData("unsplit", true)]
    [InlineData("unsplit", false)]
    [InlineData("split", true)]
    [InlineData("split", false)]
    public void TableFragmentsRecheckEachActualRunningFrame(string profile, bool occupiedFirstFrame) {
        PdfTableStyle style = ReviewTableStyle(profile);
        PdfDocument document = CreateContentDocument(content => {
            if (occupiedFirstFrame) content.Spacer(150D);
            content.Table(new[] { new[] { "ROW" } }, style: style);
        }, header => ConfigureStaticHeader(header,
            occupiedFirstFrame ? LargeBodyHeight : SmallBodyHeight, LargeBodyHeight, SmallBodyHeight));

        AssertLaterRunningFrame(document, "ROW");
    }

    [Fact]
    public void FixedRowRejectsARepeatedShortRunningFrameAfterAnOrdinaryMove() {
        PdfDocument document = CreateContentDocument(content => {
            content.Spacer(150D);
            content.Table(new[] { new[] { "ROW" } }, style: ReviewTableStyle("fixed"));
        }, header => ConfigureStaticHeader(header, LargeBodyHeight, SmallBodyHeight, SmallBodyHeight));

        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void LaterTableRowRetainsItsHeightRequirementAfterAnOrdinaryMove(bool fixedHeight) {
        PdfTableStyle style = ReviewTableStyle("fixed");
        var heights = new List<double?> { 150D, RequiredParagraphHeight };
        style.FixedRowHeights = fixedHeight ? heights : null;
        if (!fixedHeight) style.RowMinHeights = heights;
        style.MinimumBodyRowsOnFirstPage = style.MinimumBodyRowsOnLastPage = 1;
        PdfDocument document = CreateContentDocument(content =>
            content.Table(new[] { new[] { "PREFIX" }, new[] { "ROW" } }, style: style),
            header => ConfigureStaticHeader(header, LargeBodyHeight, LargeBodyHeight, SmallBodyHeight));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTROPREFIX", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Equal("ROW", pdf.GetPage(3).Text);
    }

    [Fact]
    public void OptionalRepeatedTableHeaderDoesNotRejectAnOtherwiseFittingBodyRow() {
        PdfTableStyle style = ReviewTableStyle("fixed");
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.FixedRowHeights = new List<double?> { 30D, 150D, 40D };
        style.MinimumBodyRowsOnFirstPage = style.MinimumBodyRowsOnLastPage = 1;
        PdfDocument document = CreateContentDocument(content =>
            content.Table(new[] { new[] { "HEAD" }, new[] { "PREFIX" }, new[] { "ROW" } }, style: style),
            header => ConfigureStaticHeader(header, LargeBodyHeight, SmallBodyHeight, SmallBodyHeight));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("INTROHEADPREFIX", pdf.GetPage(1).Text);
        Assert.Equal("ROW", pdf.GetPage(2).Text);
    }

    [Fact]
    public void SplittableTableRowUsesALegalShortFrameInsteadOfSearchingForTheWholeRow() {
        PdfTableStyle style = ReviewTableStyle("split");
        style.FontSize = 20D;
        style.FixedRowHeights = new List<double?> { 150D, null };
        style.MinimumBodyRowsOnFirstPage = style.MinimumBodyRowsOnLastPage = 1;
        var runs = new[] { PdfTextRun.Normal("B01\nB02\nB03\nB04") };
        // The Word adapter supplies these paragraph-level widow constraints
        // through the existing cell model. Two 25pt lines fit the short frame.
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs,
            fontSize: 20D, lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(25D)), widowControl: true) });
        PdfDocument document = CreateContentDocument(content =>
            content.Table(new[] { new[] { new PdfTableCell("PREFIX") }, new[] { cell } }, style: style),
            header => ConfigureStaticHeader(header, LargeBodyHeight, LargeBodyHeight, SmallBodyHeight));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTROPREFIX", pdf.GetPage(1).Text);
        Assert.Equal("B01B02", pdf.GetPage(2).Text);
        Assert.Equal("B03B04", pdf.GetPage(3).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ContainerKeepWithNextRechecksRunningFramesAfterAnOrdinaryMove(bool laterFrameFits) {
        PdfDocument document = CreateContentDocument(content => {
            content.Spacer(90D);
            content.Panel(panel => panel.Paragraph(paragraph => paragraph.Text("PANEL"),
                style: TargetStyle(keepTogether: false)), new PdfPanelStyle {
                    PaddingX = 0D, PaddingY = 0D, SpacingBefore = 0D, SpacingAfter = 0D,
                    KeepTogether = true, KeepWithNext = true
                });
            content.Paragraph(paragraph => paragraph.Text("NEXT"), style: TargetStyle(keepTogether: false));
        }, header => ConfigureStaticHeader(header, LargeBodyHeight,
            laterFrameFits ? LargeBodyHeight : SmallBodyHeight, SmallBodyHeight));

        if (!laterFrameFits) {
            Assert.Throws<ArgumentException>(() => document.ToBytes());
            return;
        }
        AssertLaterRunningFrame(document, "PANELNEXT");
    }

    [Theory]
    [InlineData("table", true)]
    [InlineData("table", false)]
    [InlineData("list", true)]
    [InlineData("list", false)]
    public void KeptGroupsRemainWholeAfterTheirKeepWithNextMove(string kind, bool laterFrameFits) {
        PdfDocument document = CreateContentDocument(content => {
            content.Spacer(90D);
            if (kind == "table") {
                PdfTableStyle style = ReviewTableStyle("kept");
                style.KeepWithNext = true;
                style.FixedRowHeights = new List<double?> { 40D, 40D };
                style.MinimumBodyRowsOnFirstPage = style.MinimumBodyRowsOnLastPage = 1;
                content.Table(new[] { new[] { "FIRST" }, new[] { "SECOND" } }, style: style);
            } else {
                content.Bullets(new[] { "FIRST", "SECOND" }, style: new PdfListStyle {
                    FontSize = 10D, LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(40D)), ItemSpacing = 0D,
                    SpacingBefore = 0D, SpacingAfter = 0D, KeepTogether = true, KeepWithNext = true,
                    KeepTogetherOverflow = PdfKeepTogetherOverflowBehavior.Reject
                });
            }
            // The 80pt group fits the remaining 100pt. Its 30pt successor
            // triggers the move but can itself fit the short 50pt frame.
            content.Paragraph(paragraph => paragraph.Text("NEXT"), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(30D)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), WidowControl = false
            });
        }, header => ConfigureStaticHeader(header, LargeBodyHeight,
            laterFrameFits ? LargeBodyHeight : SmallBodyHeight, SmallBodyHeight));

        if (!laterFrameFits) {
            Assert.Throws<ArgumentException>(() => document.ToBytes());
            return;
        }
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Equal(new[] { "FIRST", "SECOND", "NEXT" },
            Regex.Matches(pdf.GetPage(3).Text, "FIRST|SECOND|NEXT").Cast<Match>().Select(match => match.Value));
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("panel")]
    public void AtomicNoteMoveRechecksTheKeptBodyAgainstItsNewRunningFrame(string kind) {
        var capture = new PdfLayoutPositionCapture();
        PdfDocument document = CreateContentDocument(content => {
            content.Spacer(90D);
            PdfFootnoteDefinition note = content.DefineFootnote(body => body.Text("NOTE", style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(40D)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), WidowControl = false
            }));
            Action<PdfContentBuilder> body = group => group.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("REF").WithFootnoteReference(note)
            }), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(80D)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), WidowControl = false
            });
            if (kind == "flow") content.Flow(body, new PdfFlowOptions { KeepTogether = true }, capture);
            else content.Panel(body, new PdfPanelStyle {
                PaddingX = 0D, PaddingY = 0D, SpacingBefore = 0D, SpacingAfter = 0D, KeepTogether = true
            });
        }, header => ConfigureStaticHeader(header, LargeBodyHeight, LargeBodyHeight, SmallBodyHeight));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Single(pdf.GetPage(3).GetWords(), word => word.Text == "REF");
        Assert.Single(pdf.GetPage(3).GetWords(), word => word.Text == "NOTE");
        if (kind == "flow") {
            Assert.False(capture.WasSkipped);
            Assert.Equal(3, Assert.Single(capture.Regions).PageNumber);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void IndivisibleListLinesRecheckEachActualRunningFrame(bool occupiedFirstFrame) {
        PdfDocument document = CreateContentDocument(content => {
            if (occupiedFirstFrame) content.Spacer(150D);
            content.Bullets(new[] { "LIST" }, style: new PdfListStyle {
                FontSize = 10D, LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(RequiredParagraphHeight)),
                SpacingBefore = 0D, SpacingAfter = 0D, ItemSpacing = 0D
            });
        }, header => ConfigureStaticHeader(header,
            occupiedFirstFrame ? LargeBodyHeight : SmallBodyHeight, LargeBodyHeight, SmallBodyHeight));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Single(Regex.Matches(pdf.GetPage(3).Text, "LIST").Cast<Match>());
    }

    private static PdfTableStyle ReviewTableStyle(string profile) {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = style.CellPaddingY = 0D;
        style.SpacingBefore = style.SpacingAfter = 0D;
        style.KeepTogether = profile == "kept";
        style.AllowRowBreakAcrossPages = profile == "split";
        if (profile is "kept" or "fixed")
            style.FixedRowHeights = new List<double?> { RequiredParagraphHeight };
        else style.FontSize = RequiredParagraphHeight / 1.25D;
        return style;
    }

    private static void AssertLaterRunningFrame(PdfDocument document, string target) {
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Empty(pdf.GetPage(2).Text);
        Assert.Equal(target, pdf.GetPage(3).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstPageStoryCanYieldToALargerStaticDefaultWithoutAPageLimit(bool contextualFirst) {
        PdfDocument document = CreateContentDocument(content =>
            content.Paragraph(paragraph => paragraph.Text("TARGET"), style: TargetStyle(keepTogether: false)),
            configureHeader: header => {
                header.Content(story => story.Spacer(HeaderHeight(LargeBodyHeight)), HeaderDistance);
                ConfigureFirstPageHeader(header, contextualFirst);
            }, maximumPages: null);

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether: false, targetPage: 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstPageStoryCanYieldToAnAbsentDefaultWithoutAPageLimit(bool contextualFirst) {
        PdfDocument document = CreateContentDocument(content =>
            content.Paragraph(paragraph => paragraph.Text("TARGET"), style: TargetStyle(keepTogether: false)),
            configureHeader: header => ConfigureFirstPageHeader(header, contextualFirst), maximumPages: null);

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        AssertBodyPlacement(pdf, keepTogether: false, targetPage: 2);
    }

    private static void ConfigureFirstPageHeader(PdfHeaderBuilder header, bool contextualFirst) {
        if (contextualFirst)
            header.FirstPageContent(context => story => story.Spacer(HeaderHeight(SmallBodyHeight)), HeaderDistance);
        else header.FirstPageContent(story => story.Spacer(HeaderHeight(SmallBodyHeight)), HeaderDistance);
    }

    private static PdfDocument CreateDocument(bool keepTogether, bool explicitBreak,
        double firstBodyHeight = SmallBodyHeight, double laterBodyHeight = LargeBodyHeight,
        double? evenBodyHeight = null) => CreateContentDocument(content => {
            if (explicitBreak) content.PageBreak();
            content.Paragraph(paragraph => {
                if (keepTogether) {
                    paragraph.Text("KEEPFIRST").LineBreak()
                        .Text("KEEPMIDDLE").LineBreak().Text("KEEPLAST");
                } else paragraph.Text("TARGET");
            }, style: TargetStyle(keepTogether));
        }, header => ConfigureStaticHeader(header, firstBodyHeight, laterBodyHeight, evenBodyHeight));

    private static PdfDocument CreateContentDocument(Action<PdfContentBuilder> target,
        Action<PdfHeaderBuilder>? configureHeader = null, int? maximumPages = 4) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(240D),
            PageHeight = global::OfficeIMO.OfficeLength.Points(PageHeight),
            MarginTop = global::OfficeIMO.OfficeLength.Points(PhysicalMargin),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(PhysicalMargin),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(24D),
            MarginRight = global::OfficeIMO.OfficeLength.Points(24D),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10D),
            MaxGeneratedPages = maximumPages
        };
        return PdfDocument.Create(options).Compose(document => document.Page(page => {
            page.Header(configureHeader ?? (header => ConfigureStaticHeader(header, SmallBodyHeight, LargeBodyHeight)));
            page.Content(content => {
                content.Paragraph(paragraph => paragraph.Text("INTRO"), style: new PdfParagraphStyle {
                    LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(SmallBodyHeight / 5D)),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
                    WidowControl = false
                });
                target(content);
            });
        }));
    }

    private static PdfParagraphStyle TargetStyle(bool keepTogether) => new() {
        LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(keepTogether
            ? RequiredParagraphHeight / 3D : RequiredParagraphHeight)),
        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
        SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
        WidowControl = false,
        KeepTogether = keepTogether,
        KeepTogetherOverflow = PdfKeepTogetherOverflowBehavior.Reject
    };

    private static void ConfigureStaticHeader(PdfHeaderBuilder header, double firstBodyHeight,
        double laterBodyHeight, double? evenBodyHeight = null) {
        // Each spacer gives its variant the declared body capacity:
        // page height - edge distance - story height - bottom margin.
        header.Content(content => content.Spacer(HeaderHeight(laterBodyHeight)), HeaderDistance)
            .FirstPageContent(content => content.Spacer(HeaderHeight(firstBodyHeight)), HeaderDistance);
        if (evenBodyHeight.HasValue)
            header.EvenPagesContent(content => content.Spacer(HeaderHeight(evenBodyHeight.Value)), HeaderDistance);
    }

    private static double HeaderHeight(double bodyHeight) =>
        PageHeight - HeaderDistance - PhysicalMargin - bodyHeight;

    private static void AssertBodyPlacement(PdfPigDocument pdf, bool keepTogether, int targetPage) {
        string target = keepTogether ? "KEEPFIRSTKEEPMIDDLEKEEPLAST" : "TARGET";
        Assert.Equal(targetPage, pdf.NumberOfPages);
        if (targetPage == 1) {
            Assert.Equal("INTRO" + target, pdf.GetPage(1).Text);
        } else {
            Assert.Equal("INTRO", pdf.GetPage(1).Text);
            Assert.Equal(target, pdf.GetPage(targetPage).Text);
        }
    }
}
