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
        var style = TableStyles.Minimal();
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

        Assert.Throws<ArgumentException>(() => document.ToBytes(cancellation.Token));
    }

    [Fact]
    public void HeadingLineReachesTheLargerOddFrameAfterShortFirstAndEvenFrames() {
        PdfDocument document = CreateContentDocument(content =>
            content.Heading(1, "HEADING", style: new PdfHeadingStyle {
                FontSize = 10D,
                LineSpacing = PdfLineSpacing.Exactly(RequiredParagraphHeight),
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
                LineSpacing = PdfLineSpacing.Exactly(RequiredParagraphHeight),
                SpacingBefore = 0D,
                SpacingAfter = 0D,
                ItemSpacing = 0D,
                KeepTogether = true,
                KeepTogetherOverflow = PdfKeepTogetherOverflowBehavior.Reject
            }));

        using var pdf = PdfPigDocument.Open(document.ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("INTRO", pdf.GetPage(1).Text);
        Assert.Equal(1, Regex.Matches(pdf.GetPage(2).Text, "LIST").Count);
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
        var options = new PdfOptions {
            PageWidth = 240D,
            PageHeight = PageHeight,
            MarginTop = PhysicalMargin,
            MarginBottom = PhysicalMargin,
            MarginLeft = 24D,
            MarginRight = 24D,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10D,
            MaxGeneratedPages = maximumPages
        };
        return PdfDocument.Create(options).Compose(document => document.Page(page => {
            page.Header(configureHeader ?? (header => ConfigureStaticHeader(header, SmallBodyHeight, LargeBodyHeight)));
            page.Content(content => {
                content.Paragraph(paragraph => paragraph.Text("INTRO"), style: new PdfParagraphStyle {
                    LineSpacing = PdfLineSpacing.Exactly(SmallBodyHeight / 5D),
                    SpacingBefore = 0D,
                    SpacingAfter = 0D,
                    WidowControl = false
                });
                target(content);
            });
        }));
    }

    private static PdfParagraphStyle TargetStyle(bool keepTogether) => new() {
        LineSpacing = PdfLineSpacing.Exactly(keepTogether
            ? RequiredParagraphHeight / 3D : RequiredParagraphHeight),
        SpacingBefore = 0D,
        SpacingAfter = 0D,
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
