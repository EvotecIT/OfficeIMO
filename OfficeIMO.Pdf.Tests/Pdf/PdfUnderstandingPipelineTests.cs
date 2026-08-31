using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfUnderstandingPipelineTests {
    [Fact]
    public void FastPipeline_ExposesAllStagesAndCallerOrderedPages() {
        byte[] pdf = PdfDocument.Create()
            .H1("Pipeline heading")
            .Paragraph(p => p.Text("First body line"))
            .PageBreak()
            .Paragraph(p => p.Text("Second page body"))
            .ToBytes();

        PdfDocumentReadResult result = Read(
            pdf,
            new PdfUnderstandingPipelineOptions(),
            PdfReadProfile.Fast,
            PdfPageSelection.From(2, 1));
        PdfUnderstandingPageResult[] pages = result.Pages.Select(static page => page.Analysis).ToArray();

        Assert.Equal(new[] { 2, 1 }, result.Pages.Select(static page => page.PageNumber));
        Assert.All(pages, page => {
            Assert.NotEmpty(page.DecodedRuns);
            Assert.NotEmpty(page.Words);
            Assert.NotEmpty(page.Lines);
            Assert.NotEmpty(page.Regions);
            Assert.NotEmpty(page.ReadingOrder);
            Assert.NotEmpty(page.Elements);
            Assert.All(page.Words, word => { Assert.InRange(word.Confidence, 0D, 1D); Assert.NotEmpty(word.Evidence); });
            Assert.All(page.Lines, line => { Assert.InRange(line.Confidence, 0D, 1D); Assert.NotEmpty(line.Evidence); });
            Assert.All(page.Regions, region => { Assert.InRange(region.Confidence, 0D, 1D); Assert.NotEmpty(region.Evidence); });
            Assert.All(page.ReadingOrderEvidence, order => { Assert.InRange(order.Confidence, 0D, 1D); Assert.NotEmpty(order.Evidence); });
            Assert.All(page.Elements, element => { Assert.InRange(element.Confidence, 0D, 1D); Assert.NotEmpty(element.Evidence); });
            Assert.Equal(new[] { "glyph-decoding", "word-grouping", "line-grouping", "page-segmentation", "reading-order", "semantic-classification" }, page.Trace.Select(static trace => trace.Stage));
        });
        Assert.All(pages.SelectMany(static page => page.Trace), trace =>
            Assert.Equal(typeof(PdfAdvancedUnderstandingStages).Assembly, trace.ProviderType.Assembly));
    }

    [Fact]
    public void Pipeline_UsesCallerSuppliedStageAndRecordsItsProvider() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Top region"))
            .Paragraph(p => p.Text("Bottom region"), style: new PdfParagraphStyle { SpacingBefore = 40 })
            .ToBytes();
        var custom = new ReverseReadingOrderStage();
        var options = new PdfUnderstandingPipelineOptions { ReadingOrder = custom };

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;
        PdfUnderstandingPageResult baseline = Assert.Single(Read(pdf, new PdfUnderstandingPipelineOptions()).Pages).Analysis;

        Assert.Equal(typeof(ReverseReadingOrderStage), Assert.Single(page.Trace, static trace => trace.Stage == "reading-order").ProviderType);
        Assert.Equal(page.Regions.Reverse().Select(static region => region.Text), page.ReadingOrder.Select(static region => region.Text));
        foreach (PdfUnderstandingStageTrace trace in page.Trace.Where(static trace => trace.Stage != "reading-order")) {
            Assert.Equal(
                Assert.Single(baseline.Trace, baselineTrace => baselineTrace.Stage == trace.Stage).ProviderType,
                trace.ProviderType);
        }
    }

    [Fact]
    public void Pipeline_RejectsOversizedCustomStageArtifacts() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = new PdfUnderstandingPipelineOptions {
            GlyphDecoding = new FixedGlyphStage(new[] {
                new PdfTextSpan("one", "F1", 12, 10, 10, 20),
                new PdfTextSpan("two", "F1", 12, 40, 10, 20)
            }),
            MaxRunsPerPage = 1
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => Read(pdf, options));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void AdvancedPipeline_GroupsRotatedBaselinesAndOrdersMultipleColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left top", "F1", 12, 50, 700, 48),
            new PdfTextSpan("Left bottom", "F1", 12, 50, 650, 60),
            new PdfTextSpan("Right top", "F1", 12, 300, 700, 54),
            new PdfTextSpan("Right bottom", "F1", 12, 300, 650, 66),
            new PdfTextSpan("Vertical one", "F1", 12, 520, 300, 66, rotationDegrees: 90),
            new PdfTextSpan("Vertical two", "F1", 12, 520, 370, 66, rotationDegrees: 90)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfDocumentReadResult result = Read(pdf, options);
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        Assert.Contains(page.Lines, line => Math.Abs(line.RotationDegrees - 90D) < 0.1D && line.Text.Contains("Vertical one Vertical two", StringComparison.Ordinal));
        string[] horizontalOrder = page.ReadingOrder.Select(region => region.Text).Where(text => text.StartsWith("Left", StringComparison.Ordinal) || text.StartsWith("Right", StringComparison.Ordinal)).ToArray();
        Assert.Equal(new[] { "Left top", "Left bottom", "Right top", "Right bottom" }, horizontalOrder);
        Assert.True(result.Text.IndexOf("Left top", StringComparison.Ordinal) < result.Text.IndexOf("Left bottom", StringComparison.Ordinal));
        Assert.True(result.Text.IndexOf("Left bottom", StringComparison.Ordinal) < result.Text.IndexOf("Right top", StringComparison.Ordinal));
        Assert.True(result.Text.IndexOf("Right top", StringComparison.Ordinal) < result.Text.IndexOf("Right bottom", StringComparison.Ordinal));
        Assert.Equal(typeof(PdfAdvancedUnderstandingStages).Assembly, Assert.Single(page.Trace, trace => trace.Stage == "reading-order").ProviderType.Assembly);
    }

    [Fact]
    public void AdvancedLineGrouping_PreservesLongWordsFromOneSourceRun() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        const string text = "Sensitive payroll token PAY-SECRET-2026";
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan(text, "F1", 12D, 50D, 700D, 300D)
        });

        PdfDocumentReadResult result = Read(pdf, options);

        Assert.Equal(text, Assert.Single(Assert.Single(result.Pages).Analysis.Lines).Text);
        Assert.Contains(text, result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void AdvancedLineGrouping_JoinsAdjacentRunFragmentsAndSpacesPositiveGaps() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Hel", "F1", 12D, 50D, 700D, 18D),
            new PdfTextSpan("lo", "F1", 12D, 68D, 700D, 12D),
            new PdfTextSpan("world", "F1", 12D, 90D, 700D, 30D)
        });

        PdfDocumentReadResult result = Read(pdf, options);

        Assert.Equal("Hello world", Assert.Single(Assert.Single(result.Pages).Analysis.Lines).Text);
        Assert.Contains("Hello world", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Hel lo", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void AdvancedPipeline_UsesRecursiveWhitespaceCutsForIndentedColumnContent() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left one", "F1", 12, 50, 700, 110),
            new PdfTextSpan("Left two", "F1", 12, 80, 650, 110),
            new PdfTextSpan("Left three", "F1", 12, 50, 600, 110),
            new PdfTextSpan("Right one", "F1", 12, 320, 700, 110),
            new PdfTextSpan("Right two", "F1", 12, 320, 650, 110),
            new PdfTextSpan("Right three", "F1", 12, 320, 600, 110)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(new[] {
            "Left one", "Left two", "Left three",
            "Right one", "Right two", "Right three"
        }, page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_ForceSingleColumnOrdersRowsBeforeColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left top", "F1", 12, 50, 700, 90),
            new PdfTextSpan("Right top", "F1", 12, 320, 700, 90),
            new PdfTextSpan("Left bottom", "F1", 12, 50, 640, 90),
            new PdfTextSpan("Right bottom", "F1", 12, 320, 640, 90)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(
            pdf,
            options,
            layoutOptions: new PdfTextLayoutOptions { ForceSingleColumn = true }).Pages).Analysis;

        Assert.Equal(new[] { "Left top", "Right top", "Left bottom", "Right bottom" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_OrdersStaggeredRegionsTopToBottomInsteadOfAsColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Upper right heading", "F1", 16, 320, 700, 150),
            new PdfTextSpan("Lower left body", "F1", 11, 50, 500, 120)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(new[] { "Upper right heading", "Lower left body" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdvancedPipeline_OrdersInterleavedStaggeredRegionsTopToBottom(bool mirrorColumns) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        double outerColumn = mirrorColumns ? 320 : 50;
        double middleColumn = mirrorColumns ? 50 : 320;
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Upper outer", "F1", 12, outerColumn, 700, 110),
            new PdfTextSpan("Middle inner", "F1", 12, middleColumn, 600, 110),
            new PdfTextSpan("Lower outer", "F1", 12, outerColumn, 500, 110)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(new[] { "Upper outer", "Middle inner", "Lower outer" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_RequiresActualBandOverlapBeforeSelectingColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Upper left", "F1", 12, 50, 700, 110),
            new PdfTextSpan("Upper middle right", "F1", 12, 320, 600, 130),
            new PdfTextSpan("Lower middle left", "F1", 12, 50, 500, 130),
            new PdfTextSpan("Lower right", "F1", 12, 320, 400, 110)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(new[] {
            "Upper left", "Upper middle right", "Lower middle left", "Lower right"
        }, page.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedPipeline_LeavesUniqueEdgesAndUnvalidatedColumnsAsContent() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("Quarterly report", "F1", 12, 50, 800, 90),
            new PdfTextSpan("Item", "F1", 11, 50, 500, 24), new PdfTextSpan("Amount", "F1", 11, 90, 500, 42),
            new PdfTextSpan("Licenses", "F1", 11, 50, 482, 24), new PdfTextSpan("42", "F1", 11, 90, 482, 12),
            new PdfTextSpan("Figure 1. Revenue by region", "F1", 10, 50, 400, 150),
            new PdfTextSpan("1 Audited values exclude pending adjustments.", "F1", 8, 50, 20, 190)
        });
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.DoesNotContain(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Header);
        Assert.DoesNotContain(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Footer);
        Assert.DoesNotContain(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Table);
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Paragraph && element.Region.Text == "Quarterly report");
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Paragraph && element.Region.Text.Contains("Item Amount", StringComparison.Ordinal));
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Caption);
        Assert.Contains(page.Elements, element => element.Kind == PdfUnderstandingSemanticKind.Footnote);
    }

    [Fact]
    public void AdvancedSemanticClassification_UsesVisualCoordinatesForRotatedFootnotes() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        pdf = PdfPageEditor.SetCropBox(pdf, 100D, 0D, 500D, 600D);
        pdf = PdfPageEditor.RotatePages(pdf, 90);
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(sourcePage, 1, new PdfTextLayoutOptions(), 10_000, 10_000);
        PdfUnderstandingRegion[] regions = {
            CreateUnderstandingRegion("Body one", 300D, 300D, 12D),
            CreateUnderstandingRegion("Body two", 300D, 250D, 12D),
            CreateUnderstandingRegion("Body three", 300D, 200D, 12D),
            CreateUnderstandingRegion("Visual footnote", 130D, 300D, 8D),
            CreateUnderstandingRegion("Raw bottom only", 300D, 20D, 8D)
        };

        IReadOnlyList<PdfUnderstandingSemanticElement> elements =
            PdfAdvancedUnderstandingStages.SemanticClassification.Classify(context, regions);

        Assert.Equal(PdfUnderstandingSemanticKind.Footnote,
            Assert.Single(elements, static element => element.Region.Text == "Visual footnote").Kind);
        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(elements, static element => element.Region.Text == "Raw bottom only").Kind);
    }

    [Fact]
    public void Sections_KeepTableLinesOutOfUngroupedTextBlocks() {
        byte[] pdf = PdfDocument.Create()
            .H1("Table section")
            .Paragraph(paragraph => paragraph.Text("Section introduction."))
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Quality", "Premium" }
            })
            .ToBytes();

        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.SemanticClassification = new TableSectionClassificationStage();
        PdfDocumentReadResult result = Read(pdf, options);
        PdfLogicalSection section = Assert.Single(result.Sections);

        Assert.Single(section.Tables);
        Assert.DoesNotContain(section.TextBlocks, block =>
            block.Text.Contains("Metric", StringComparison.Ordinal) ||
            block.Text.Contains("Quality", StringComparison.Ordinal));
        Assert.Contains("Metric\tValue", result.Text, StringComparison.Ordinal);
        Assert.Contains("Quality\tPremium", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void LogicalProjection_PreservesRotatedLineVisualBounds() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Vertical text", "F1", 12D, 220D, 300D, 72D, rotationDegrees: 90D)
        });

        PdfLogicalPage page = Assert.Single(Read(pdf, options).Pages);
        PdfLogicalTextBlock block = Assert.Single(page.TextBlocks);
        PdfLogicalVisualBounds bounds = Assert.IsType<PdfLogicalVisualBounds>(block.VisualBounds);
        PdfLogicalReadingOrderItem item = Assert.Single(PdfLogicalReadingOrderAnalysis.Analyze(page));

        Assert.True(bounds.Height > 60D);
        Assert.True(bounds.Width > 5D);
        Assert.True(item.HasGeometry);
        Assert.True(item.Bottom - item.Top > 60D);
    }

    [Fact]
    public void FastAndStructured_UseTheSameCanonicalPageStagesAndArtifacts() {
        byte[] pdf = PdfDocument.Create()
            .H1("Shared page pipeline")
            .Paragraph(paragraph => paragraph.Text("Body content before the table."))
            .Table(new[] {
                new[] { "Account", "Amount" },
                new[] { "North", "42" }
            })
            .ToBytes();

        PdfDocumentReadResult fast = PdfDocument.Load(pdf).Read(new PdfReadOptions { Profile = PdfReadProfile.Fast });
        PdfDocumentReadResult structured = PdfDocument.Load(pdf).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured });
        PdfUnderstandingPageResult fastPage = Assert.Single(fast.Pages).Analysis;
        PdfUnderstandingPageResult structuredPage = Assert.Single(structured.Pages).Analysis;

        Assert.Equal(fastPage.Trace.Select(static trace => trace.ProviderType), structuredPage.Trace.Select(static trace => trace.ProviderType));
        Assert.Equal(fastPage.DecodedRuns.Select(static run => run.Text), structuredPage.DecodedRuns.Select(static run => run.Text));
        Assert.Equal(fastPage.Words.Select(static word => word.Text), structuredPage.Words.Select(static word => word.Text));
        Assert.Equal(fastPage.Lines.Select(static line => line.Text), structuredPage.Lines.Select(static line => line.Text));
        Assert.Equal(fastPage.Regions.Select(static region => region.Text), structuredPage.Regions.Select(static region => region.Text));
        Assert.Equal(fastPage.ReadingOrder.Select(static region => region.Text), structuredPage.ReadingOrder.Select(static region => region.Text));
    }

    [Fact]
    public void FastRead_LeavesDocumentWideHeadingTiersToStructuredProfile() {
        byte[] pdf = PdfDocument.Create()
            .H1("Larger page heading")
            .PageBreak()
            .H2("Smaller page heading")
            .ToBytes();
        var pipeline = new PdfUnderstandingPipelineOptions {
            SemanticClassification = new HeadingClassificationStage()
        };

        PdfDocumentReadResult fast = Read(pdf, pipeline, PdfReadProfile.Fast);
        PdfDocumentReadResult structured = Read(pdf, pipeline, PdfReadProfile.Structured);

        Assert.Equal(new[] { 1, 1 }, fast.Pages.SelectMany(static page => page.Headings).Select(static heading => heading.Level));
        Assert.Equal(new[] { 1, 2 }, structured.Pages.SelectMany(static page => page.Headings).Select(static heading => heading.Level));
    }

    [Fact]
    public void StructuredRead_PromotesPageLocalHeadingToItsDocumentFontTier() {
        byte[] pdf = PdfDocument.Create()
            .H1("Only document heading")
            .Paragraph(paragraph => paragraph.Text("Document body."))
            .ToBytes();
        var pipeline = new PdfUnderstandingPipelineOptions {
            SemanticClassification = new HeadingLevelTwoClassificationStage()
        };

        PdfLogicalHeading fast = Assert.Single(Read(pdf, pipeline, PdfReadProfile.Fast).Headings);
        PdfLogicalHeading structured = Assert.Single(Read(pdf, pipeline, PdfReadProfile.Structured).Headings);

        Assert.Equal(2, fast.Level);
        Assert.Equal(1, structured.Level);
    }

    [Fact]
    public void StructuredRead_PreservesDigitsWhenMatchingOutlineTitles() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateNumberedOutlinePdf()).Read();
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        PdfUnderstandingSemanticElement matched = Assert.Single(page.Elements, element =>
            element.Region.Text == "Section 2" &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.outline-heading"));
        Assert.Equal(PdfUnderstandingSemanticKind.Heading, matched.Kind);
        Assert.DoesNotContain(page.Elements, element =>
            element.Region.Text == "Section 1" &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.outline-heading"));
    }

    [Fact]
    public void StructuredRead_ScopesTaggedMcidsToTheirContentStreams() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateScopedTaggedMcidPdf()).Read();
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        AssertTaggedHeading(page, "Page heading", 1);
        Assert.Contains(page.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Paragraph &&
            element.Region.Text == "Form paragraph" &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.tagged-pdf-role"));
        Assert.DoesNotContain(page.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Heading &&
            element.Region.Text == "Form paragraph");

        PdfTextSpan pageRun = Assert.Single(page.DecodedRuns, static run => run.Text == "Page heading");
        PdfTextSpan formRun = Assert.Single(page.DecodedRuns, static run => run.Text == "Form paragraph");
        Assert.Null(pageRun.ContentStreamObjectNumber);
        Assert.Equal(6, formRun.ContentStreamObjectNumber);
        PdfMarkedContentReference formReference = Assert.Single(
            result.TaggedContent!.StructureElements,
            static element => element.StructureType == "P").MarkedContentReferences.Single();
        Assert.Equal(6, formReference.ContentStreamObjectNumber);
    }

    [Fact]
    public void DocumentEnrichment_EnforcesSemanticElementLimitAfterOutlineProjection() {
        PdfReadDocument document = PdfReadDocument.Open(CreateNumberedOutlinePdf());
        PdfUnderstandingLine first = CreateUnderstandingLine("Section 1", 72D, 150D, 700D);
        PdfUnderstandingLine second = CreateUnderstandingLine("Section 2", 72D, 150D, 650D);
        var region = new PdfUnderstandingRegion(new[] { first, second });
        var page = new PdfUnderstandingPageResult(
            1,
            first.Words.SelectMany(static word => word.SourceRuns)
                .Concat(second.Words.SelectMany(static word => word.SourceRuns))
                .ToArray(),
            first.Words.Concat(second.Words).ToArray(),
            new[] { first, second },
            new[] { region },
            new[] { region },
            Array.Empty<PdfReadingOrderEvidence>(),
            new[] {
                new PdfUnderstandingSemanticElement(
                    region,
                    PdfUnderstandingSemanticKind.Paragraph,
                    0.8D)
            },
            Array.Empty<PdfUnderstandingStageTrace>());

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocumentSemanticEnricher.Enrich(
                document,
                new[] { 1 },
                new[] { page },
                1,
                10_000,
                CancellationToken.None));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void LogicalProjection_ReusesCanonicalDecodedRunsAndHonorsCancellation() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Source content must not be decoded twice"))
            .ToBytes();
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        PdfUnderstandingLine line = CreateUnderstandingLine("Canonical decoded content", 72D, 250D, 700D);
        PdfTextSpan[] runs = line.Words.SelectMany(static word => word.SourceRuns).ToArray();
        var region = new PdfUnderstandingRegion(new[] { line });
        var analysis = new PdfUnderstandingPageResult(
            1,
            runs,
            line.Words,
            new[] { line },
            new[] { region },
            new[] { region },
            Array.Empty<PdfReadingOrderEvidence>(),
            new[] { new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Paragraph, 0.8D) },
            Array.Empty<PdfUnderstandingStageTrace>());

        PdfLogicalPage projected = PdfLogicalPage.From(
            document,
            document.Pages[0],
            1,
            new PdfTextLayoutOptions(),
            analysis: analysis,
            cancellationToken: CancellationToken.None);

        Assert.Contains(projected.TextBlocks, static block => block.Text == "Canonical decoded content");
        Assert.DoesNotContain(projected.TextBlocks, static block => block.Text.Contains("Source content", StringComparison.Ordinal));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => PdfLogicalPage.From(
            document,
            document.Pages[0],
            1,
            new PdfTextLayoutOptions(),
            analysis: analysis,
            cancellationToken: cancellation.Token));
    }

    [Fact]
    public void LogicalProjection_UsesConfiguredLineArtifactsForSemanticBlocks() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Source content must not define projection lines"))
            .ToBytes();
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        var firstRun = new PdfTextSpan("Combined", "Helvetica", 20D, 72D, 700D, 90D);
        var secondRun = new PdfTextSpan("heading", "Helvetica", 20D, 72D, 650D, 80D);
        var line = new PdfUnderstandingLine(new[] {
            new PdfUnderstandingWord("Combined", 72D, 162D, 700D, 20D, 0D, new[] { firstRun }),
            new PdfUnderstandingWord("heading", 72D, 152D, 650D, 20D, 0D, new[] { secondRun })
        });
        var region = new PdfUnderstandingRegion(new[] { line });
        var analysis = new PdfUnderstandingPageResult(
            1,
            new[] { firstRun, secondRun },
            line.Words,
            new[] { line },
            new[] { region },
            new[] { region },
            Array.Empty<PdfReadingOrderEvidence>(),
            new[] { new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Heading, 0.95D, level: 2) },
            Array.Empty<PdfUnderstandingStageTrace>());

        PdfLogicalPage projected = PdfLogicalPage.From(
            document,
            document.Pages[0],
            1,
            new PdfTextLayoutOptions(),
            analysis: analysis,
            cancellationToken: CancellationToken.None);

        PdfLogicalTextBlock block = Assert.Single(projected.TextBlocks);
        Assert.Equal("Combined heading", block.Text);
        PdfLogicalHeading heading = Assert.Single(projected.Headings);
        Assert.Same(block, heading.Line);
        Assert.Equal(2, heading.Level);
    }

    [Fact]
    public void LogicalProjection_ClearsDecodedTextWhenConfiguredLineStageReturnsNoLines() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Filtered source content"))
            .ToBytes();
        var pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.LineGrouping = new FixedLineGroupingStage(Array.Empty<PdfUnderstandingLine>());

        PdfDocumentReadResult result = Read(pdf, pipeline);
        PdfLogicalPage page = Assert.Single(result.Pages);

        Assert.Empty(page.Analysis.Lines);
        Assert.Empty(page.TextBlocks);
        Assert.Empty(page.Paragraphs);
        Assert.Equal(string.Empty, result.Text);
    }

    [Fact]
    public void LogicalProjection_ClaimsCombinedCustomLinesOnlyOnceAcrossSourceParagraphs() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First source paragraph"))
            .Spacer(80D)
            .Paragraph(paragraph => paragraph.Text("Second source paragraph"))
            .ToBytes();
        var pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.LineGrouping = new CombinedLineGroupingStage();

        PdfLogicalPage page = Assert.Single(Read(pdf, pipeline).Pages);
        PdfLogicalTextBlock block = Assert.Single(page.TextBlocks);
        PdfLogicalParagraph paragraph = Assert.Single(page.Paragraphs);

        Assert.Same(block, Assert.Single(paragraph.Lines));
        Assert.Contains("First source paragraph", block.Text, StringComparison.Ordinal);
        Assert.Contains("Second source paragraph", block.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void StructuredRead_InheritsSemanticRolesFromTaggedAncestors() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateNestedTaggedHeadingPdf()).Read();
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        AssertTaggedHeading(page, "Nested tagged heading", 1);
        PdfStructureElementInfo span = Assert.Single(
            result.TaggedContent!.StructureElements,
            static element => element.StructureType == "Span");
        Assert.NotEmpty(span.MarkedContentReferences);
        Assert.Equal(
            Assert.Single(result.TaggedContent.StructureElements, static element => element.StructureType == "H1").ObjectNumber,
            span.ParentObjectNumber);
    }

    [Fact]
    public void StructuredRead_StopsTaggedAncestorCyclesWithoutLeakingARole() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateCyclicTaggedStructurePdf()).Read();
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        PdfUnderstandingSemanticElement element = Assert.Single(
            page.Elements,
            static candidate => candidate.Region.Text.Contains("Cyclic tagged content", StringComparison.Ordinal));
        Assert.NotEqual(PdfUnderstandingSemanticKind.Heading, element.Kind);
        Assert.DoesNotContain(element.Evidence, static evidence => evidence.Code == "semantic.tagged-pdf-role");
    }

    [Fact]
    public void Pipeline_WorkBudgetAcceptsExactLimitAndRejectsNextUnit() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfDocumentReadResult accepted = Read(pdf, CreatePassThroughPipeline(new BudgetChargingGlyphStage(10), 10));
        Assert.Single(accepted.Pages);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            Read(pdf, CreatePassThroughPipeline(new BudgetChargingGlyphStage(11), 10)));
        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(10, exception.Limit);
        Assert.Equal(11, exception.Actual);
    }

    [Fact]
    public void Pipeline_ObservesCancellationInsideAStage() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        using var cancellation = new CancellationTokenSource();
        PdfUnderstandingPipelineOptions options = CreatePassThroughPipeline(
            new CancellingGlyphStage(cancellation, cancelAfterWorkUnit: 4),
            100);

        Assert.Throws<OperationCanceledException>(() =>
            Read(pdf, options, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void AdvancedSegmentation_ChargesCanonicalLayoutRecoveryBeforeTableProjection() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(
            sourcePage,
            1,
            new PdfTextLayoutOptions(),
            10_000,
            10_000,
            maxWorkUnitsPerPage: 1);
        PdfUnderstandingLine first = CreateGappedUnderstandingLine("Metric", "Value", 500D);
        PdfUnderstandingLine second = CreateGappedUnderstandingLine("Quality", "Premium", 480D);
        context.DecodedRuns = first.Words.SelectMany(static word => word.SourceRuns)
            .Concat(second.Words.SelectMany(static word => word.SourceRuns))
            .ToArray();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfAdvancedUnderstandingStages.PageSegmentation.Segment(context, new[] { first, second }));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void Pipeline_RemovesIgnoredPageBandsBeforeSemanticStages() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Ignored header evidence", "F1", 8D, 50D, 760D, 140D),
            new PdfTextSpan("Retained body evidence", "F1", 12D, 50D, 500D, 140D),
            new PdfTextSpan("Ignored footer evidence", "F1", 8D, 50D, 20D, 140D)
        });

        PdfLogicalPage page = Assert.Single(Read(
            pdf,
            options,
            layoutOptions: new PdfTextLayoutOptions {
                IgnoreHeaderHeight = 80D,
                IgnoreFooterHeight = 80D
            }).Pages);

        Assert.Equal("Retained body evidence", Assert.Single(page.Analysis.DecodedRuns).Text);
        Assert.DoesNotContain(page.Analysis.Elements, static element => element.Region.Text.Contains("Ignored", StringComparison.Ordinal));
        Assert.DoesNotContain(page.TextBlocks, static block => block.Text.Contains("Ignored", StringComparison.Ordinal));
    }

    [Fact]
    public void Pipeline_RemovesIgnoredPageBandsInRotatedCropCoordinates() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        pdf = PdfPageEditor.SetCropBox(pdf, 100D, 200D, 500D, 800D);
        pdf = PdfPageEditor.RotatePages(pdf, 90);
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Ignored visual header", "F1", 8D, 470D, 400D, 20D),
            new PdfTextSpan("Retained visual body", "F1", 12D, 300D, 400D, 20D),
            new PdfTextSpan("Ignored visual footer", "F1", 8D, 130D, 400D, 20D)
        });

        PdfLogicalPage page = Assert.Single(Read(
            pdf,
            options,
            layoutOptions: new PdfTextLayoutOptions {
                IgnoreHeaderHeight = 60D,
                IgnoreFooterHeight = 60D
            }).Pages);

        Assert.Equal("Retained visual body", Assert.Single(page.Analysis.DecodedRuns).Text);
        Assert.DoesNotContain(page.Analysis.Elements, static element => element.Region.Text.Contains("Ignored", StringComparison.Ordinal));
        Assert.DoesNotContain(page.TextBlocks, static block => block.Text.Contains("Ignored", StringComparison.Ordinal));
    }

    [Fact]
    public void AdvancedLineGrouping_ChargesOrderingWorkBeforeGroupingTraversal() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(
            sourcePage,
            1,
            new PdfTextLayoutOptions(),
            10_000,
            10_000,
            maxWorkUnitsPerPage: 2);
        PdfUnderstandingWord[] words = {
            CreateUnderstandingWord("second", 100D, 700D),
            CreateUnderstandingWord("first", 50D, 700D)
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfAdvancedUnderstandingStages.LineGrouping.GroupLines(context, words));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    [Fact]
    public void AdvancedSegmentation_SortsCallerSuppliedLinesBeforeVerticalGapExit() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        var options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("placeholder", "F1", 12D, 10D, 10D, 30D)
        });
        options.LineGrouping = new FixedLineGroupingStage(new[] {
            CreateUnderstandingLine("High", 50D, 110D, 700D),
            CreateUnderstandingLine("Unrelated low", 300D, 400D, 100D),
            CreateUnderstandingLine("Nearby", 52D, 112D, 684D)
        });

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;
        PdfUnderstandingRegion highRegion = Assert.Single(
            page.Regions,
            static region => region.Text.Contains("High", StringComparison.Ordinal));

        Assert.Contains("Nearby", highRegion.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Unrelated low", highRegion.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void StructuredRead_BoundsDocumentWideComparisonWork() {
        byte[] pdf = PdfDocument.Create()
            .Header(header => header.Text("Repeated document header"))
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes();
        var options = new PdfUnderstandingPipelineOptions { MaxDocumentWorkUnits = 1 };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => Read(pdf, options));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void StructuredRead_ChargesTaggedStructureIndexAgainstDocumentBudget() {
        PdfReadDocument document = PdfReadDocument.Open(CreateEmptyTaggedStructurePdf());
        Assert.Single(document.TaggedContent!.StructureElements);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocumentSemanticEnricher.Enrich(
                document,
                new[] { 1 },
                new[] { PdfUnderstandingPageResult.Empty(1) },
                10,
                1,
                CancellationToken.None));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void Pipeline_RejectsRawGlyphOverflowBeforeIgnoredPageBandsAreRemoved() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = new PdfUnderstandingPipelineOptions {
            GlyphDecoding = new FixedGlyphStage(new[] {
                new PdfTextSpan("header one", "F1", 12, 10, 790, 60),
                new PdfTextSpan("header two", "F1", 12, 80, 790, 60),
                new PdfTextSpan("body", "F1", 12, 10, 400, 30)
            }),
            MaxRunsPerPage = 2
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => Read(
            pdf,
            options,
            layoutOptions: new PdfTextLayoutOptions { IgnoreHeaderHeight = 20D }));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    [Fact]
    public void IgnoredPageBandFiltering_ObservesCancellationDuringTraversal() {
        IReadOnlyList<PdfTextSpan> spans = Enumerable.Range(0, 20)
            .Select(index => new PdfTextSpan("header", "F1", 12D, index * 10D, 790D, 30D))
            .ToArray();
        using var cancellation = new CancellationTokenSource();
        int polls = 0;

        Assert.Throws<OperationCanceledException>(() => TextLayoutEngine.FilterIgnoredPageBands(
            spans,
            PdfReadDocument.Open(PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes()).Pages[0],
            new PdfTextLayoutOptions { IgnoreHeaderHeight = 20D },
            consumeWork: null,
            cancellationCheck: () => {
                if (++polls == 5) cancellation.Cancel();
                cancellation.Token.ThrowIfCancellationRequested();
            }));

        Assert.Equal(5, polls);
    }

    [Fact]
    public void OutlineIndexing_ChargesEveryNodeBeforePageLookup() {
        PdfOutlineItem[] outlines = Enumerable.Range(0, 32)
            .Select(index => new PdfOutlineItem(
                "Off-page outline " + index,
                1,
                pageNumber: 2,
                destinationTop: null,
                isExpanded: false,
                Array.Empty<PdfOutlineItem>()))
            .ToArray();
        var budget = new PdfUnderstandingWorkBudget(8, CancellationToken.None);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocumentSemanticEnricher.IndexOutlinesByPage(outlines, budget));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(8, exception.Limit);
        Assert.Equal(9, exception.Actual);
    }

    [Fact]
    public void OutlineIndexing_ObservesCancellationAcrossOffPageNodes() {
        PdfOutlineItem[] outlines = Enumerable.Range(0, 32)
            .Select(index => new PdfOutlineItem(
                "Off-page outline " + index,
                1,
                pageNumber: 2,
                destinationTop: null,
                isExpanded: false,
                Array.Empty<PdfOutlineItem>()))
            .ToArray();
        using var cancellation = new CancellationTokenSource();
        var budget = new PdfUnderstandingWorkBudget(100, cancellation.Token);
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            PdfDocumentSemanticEnricher.IndexOutlinesByPage(outlines, budget));
    }

    [Fact]
    public void HeadingFontTierLookup_ChargesSortingWork() {
        double[] fontSizes = Enumerable.Range(0, 64).Select(static index => 100D - index).ToArray();
        var budget = new PdfUnderstandingWorkBudget(64, CancellationToken.None);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfHeadingFontTierAnalysis.BuildLookup(fontSizes, () => budget.Consume()));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(64, exception.Limit);
        Assert.Equal(65, exception.Actual);
    }

    [Fact]
    public void HeadingFontTierLookup_ObservesCancellationDuringSorting() {
        double[] fontSizes = Enumerable.Range(0, 128).Select(static index => 200D - index).ToArray();
        using var cancellation = new CancellationTokenSource();
        int polls = 0;

        Assert.Throws<OperationCanceledException>(() => PdfHeadingFontTierAnalysis.BuildLookup(fontSizes, () => {
            if (++polls == 140) cancellation.Cancel();
            cancellation.Token.ThrowIfCancellationRequested();
        }));

        Assert.Equal(140, polls);
    }

    [Fact]
    public void StructuredRead_BindsMixedTaggedHeadingLevelsToTheirOwnMarkedContent() {
        byte[] pdf = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .H1("Tagged first level")
            .H2("Tagged second level")
            .H3("Tagged third level")
            .Paragraph(paragraph => paragraph.Text("Tagged body paragraph."))
            .ToBytes();

        PdfDocumentReadResult result = PdfDocument.Load(pdf).Read();
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        AssertTaggedHeading(page, "Tagged first level", 1);
        AssertTaggedHeading(page, "Tagged second level", 2);
        AssertTaggedHeading(page, "Tagged third level", 3);
        Assert.Contains(page.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Paragraph &&
            element.Region.Text.Contains("Tagged body paragraph", StringComparison.Ordinal) &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.tagged-pdf-role"));
        Assert.All(
            page.DecodedRuns.Where(static run => run.Text.Contains("Tagged", StringComparison.Ordinal)),
            static run => Assert.True(run.MarkedContentId.HasValue));
        Assert.Contains(result.TaggedContent!.StructureElements, element =>
            element.StructureType == "H2" && element.MarkedContentReferences.Count > 0);
    }

    [Fact]
    public void StructuredRead_IgnoresInvalidTaggedHeadingLevels() {
        byte[] pdf = CreateInvalidTaggedHeadingPdf();

        PdfDocumentReadResult result = PdfDocument.Load(pdf).Read();
        PdfUnderstandingPageResult page = Assert.Single(result.Pages).Analysis;

        Assert.Contains(page.Elements, element =>
            element.Region.Text.Contains("Malformed tagged heading", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Heading && element.Level is <= 0);
        Assert.Contains(result.TaggedContent!.StructureElements, element => element.StructureType == "H0");
    }

    [Theory]
    [InlineData("H0")]
    [InlineData("H7")]
    [InlineData("H9")]
    [InlineData("H10")]
    [InlineData("H٢")]
    [InlineData("Heading2")]
    public void TaggedHeadingLevels_RejectMalformedRoles(string role) {
        Assert.Null(PdfDocumentSemanticEnricher.HeadingLevel(role));
    }

    [Fact]
    public void SemanticClassifier_DoesNotCreateASecondTableFromUnrelatedGaps() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(sourcePage, 1, new PdfTextLayoutOptions(), 10_000, 10_000);
        var region = new PdfUnderstandingRegion(new[] {
            CreateGappedUnderstandingLine("Customer", "Pending", 500D),
            CreateGappedUnderstandingLine("Reference", "Complete", 480D)
        });

        PdfUnderstandingSemanticElement result = Assert.Single(
            PdfAdvancedUnderstandingStages.SemanticClassification.Classify(context, new[] { region }));

        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph, result.Kind);
    }

    [Fact]
    public void StructuredRead_ProjectsRepeatedPageEdgesIntoTheLogicalModel() {
        byte[] pdf = PdfDocument.Create()
            .Header(header => header.AlignLeft().Text("Quarterly operations {page}/{pages}"))
            .Footer(footer => footer.AlignCenter().Text("Confidential page {page}"))
            .Paragraph(paragraph => paragraph.Text("First page body."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page body."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Third page body."))
            .ToBytes();

        PdfDocumentReadResult result = Read(pdf, PdfUnderstandingPipelineOptions.Structured());

        Assert.Equal(3, result.Pages.Count);
        Assert.All(result.Pages, page => {
            Assert.Contains(page.Headers, block => block.Text.Contains("Quarterly operations", StringComparison.Ordinal));
            Assert.Contains(page.Footers, block => block.Text.Contains("Confidential page", StringComparison.Ordinal));
            Assert.DoesNotContain(page.Paragraphs, paragraph => paragraph.Text.Contains("Quarterly operations", StringComparison.Ordinal));
            Assert.DoesNotContain(page.Paragraphs, paragraph => paragraph.Text.Contains("Confidential page", StringComparison.Ordinal));
        });
        Assert.All(
            result.Pages.SelectMany(static page => page.Headers),
            block => Assert.Equal(PdfLogicalElementKind.Header, block.Kind));
        Assert.All(
            result.Pages.SelectMany(static page => page.Footers),
            block => Assert.Equal(PdfLogicalElementKind.Footer, block.Kind));
    }

    [Fact]
    public void StructuredRead_ReclassifiesRepeatedLargeFontPageEdgesAsHeaders() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("first placeholder"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("second placeholder"))
            .ToBytes();
        var pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Large running header", "F1", 18D, 36D, 780D, 180D),
            new PdfTextSpan("First body line", "F1", 10D, 36D, 400D, 100D),
            new PdfTextSpan("Second body line", "F1", 10D, 36D, 350D, 110D),
            new PdfTextSpan("Third body line", "F1", 10D, 36D, 300D, 100D)
        });

        PdfDocumentReadResult result = Read(pdf, pipeline);

        Assert.Equal(2, result.Pages.Count);
        Assert.All(result.Pages, page => {
            Assert.Contains(page.Headers, block => block.Text == "Large running header");
            Assert.DoesNotContain(page.Headings, heading => heading.Text == "Large running header");
            PdfUnderstandingSemanticElement runningHeader = Assert.Single(page.Analysis.Elements, element =>
                element.Kind == PdfUnderstandingSemanticKind.Header &&
                element.Region.Text == "Large running header");
            Assert.Contains(runningHeader.Evidence, evidence => evidence.Code == "semantic.large-font");
            Assert.Contains(runningHeader.Evidence, evidence => evidence.Code == "semantic.repeated-header");
        });
        Assert.DoesNotContain(result.Sections, section => section.Title == "Large running header");
    }

    [Fact]
    public void StructuredRead_KeepsDistinctNumberedEdgeHeadingsAsSections() {
        byte[] pdf = PdfDocument.Create()
            .H1("Chapter 1")
            .PageBreak()
            .H1("Chapter 2")
            .ToBytes();
        var pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.SemanticClassification = new HeadingClassificationStage();

        PdfDocumentReadResult result = Read(pdf, pipeline);

        Assert.Equal(new[] { "Chapter 1", "Chapter 2" }, result.Headings.Select(static heading => heading.Text));
        Assert.DoesNotContain(result.Pages.SelectMany(static page => page.Headers), block =>
            block.Text.StartsWith("Chapter ", StringComparison.Ordinal));
        Assert.Equal(new[] { "Chapter 1", "Chapter 2" }, result.Sections.Select(static section => section.Title));
    }

    [Fact]
    public void StructuredRead_ClassifiesRepeatedEdgesInRotatedCropCoordinates() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("first placeholder"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("second placeholder"))
            .ToBytes();
        pdf = PdfPageEditor.SetCropBox(pdf, 100D, 200D, 500D, 800D);
        pdf = PdfPageEditor.RotatePages(pdf, 90);
        var pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Visual header", "F1", 12D, 470D, 400D, 20D)
        });

        PdfDocumentReadResult result = Read(pdf, pipeline);

        Assert.Equal(2, result.Pages.Count);
        Assert.All(result.Pages, page =>
            Assert.Contains(page.Analysis.Elements, element =>
                element.Kind == PdfUnderstandingSemanticKind.Header &&
                element.Region.Text == "Visual header"));
    }

    [Fact]
    public void Sections_TrackLastPageInCallerSelectedReadingOrder() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Final selected page body"))
            .PageBreak()
            .H1("Section begins on source page two")
            .Paragraph(paragraph => paragraph.Text("Opening body"))
            .ToBytes();

        PdfDocumentReadResult result = PdfDocument.Load(pdf).Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(2, 1)
        });
        PdfLogicalSection section = Assert.Single(result.Sections);

        Assert.Equal(2, section.FirstPageNumber);
        Assert.Equal(1, section.LastPageNumber);
        Assert.Contains(section.Paragraphs, paragraph =>
            paragraph.Text.Contains("Final selected page body", StringComparison.Ordinal));
    }

    [Fact]
    public void AdvancedPipeline_UsesCanonicalLayoutTableRegions() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Table metric", "Value" },
                new[] { "Quality", "Premium" }
            })
            .ToBytes();

        PdfUnderstandingPageResult page = Assert.Single(Read(
            pdf,
            PdfUnderstandingPipelineOptions.Structured()).Pages).Analysis;

        PdfUnderstandingSemanticElement table = Assert.Single(
            page.Elements,
            static element => element.Kind == PdfUnderstandingSemanticKind.Table);
        Assert.Contains("Table metric", table.Region.Text, StringComparison.Ordinal);
        Assert.Contains("Premium", table.Region.Text, StringComparison.Ordinal);
        Assert.Contains(table.Region.Evidence, static evidence => evidence.Code == "region.canonical-table");
        Assert.DoesNotContain(
            page.Elements,
            static element => element.Kind == PdfUnderstandingSemanticKind.Paragraph &&
                              element.Region.Text.Contains("Premium", StringComparison.Ordinal));
    }

    [Fact]
    public void LogicalProjection_PreservesAuthoritativeUnmarkedListItems() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Semantically tagged list entry"))
            .ToBytes();
        var options = new PdfUnderstandingPipelineOptions {
            SemanticClassification = new UnmarkedListClassificationStage()
        };

        PdfLogicalPage page = Assert.Single(Read(pdf, options).Pages);
        PdfLogicalListItem item = Assert.Single(page.ListItems);

        Assert.Equal(PdfLogicalElementKind.ListItem, item.Line.Kind);
        Assert.Equal("Semantically tagged list entry", item.Text);
        Assert.Equal(string.Empty, item.Marker);
        Assert.Equal(2, item.Level);
        Assert.Equal(0.95D, item.Confidence, 3);
        Assert.Contains(item.Evidence, static evidence => evidence.Code == "semantic.test-list-item");
    }

    [Fact]
    public void LogicalProjection_KeepsMultiLineTaggedListItemAsOneReaderBlock() {
        var options = new PdfUnderstandingPipelineOptions {
            PageSegmentation = new HeadingThenRemainingRegionStage(),
            ReadingOrder = new IdentityReadingOrderStage(),
            SemanticClassification = new ParagraphClassificationStage()
        };

        PdfDocumentReadResult logical = Read(CreateMultiLineTaggedListItemPdf(), options);
        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfLogicalListItem item = Assert.Single(page.ListItems);

        Assert.Equal(2, item.Lines.Count);
        Assert.Same(item.Lines[0], item.Line);
        Assert.Equal("First tagged list line continuation tagged list line", item.Text);
        Assert.Equal(item.Text, string.Concat(item.Runs.Select(static run => run.Text)));
        Assert.Contains(item.Evidence, static evidence => evidence.Code == "semantic.tagged-pdf-role");
        Assert.Single(PdfLogicalReadingOrderAnalysis.Analyze(page),
            static candidate => candidate.Kind == PdfLogicalReadingOrderKind.ListItem);
        PdfLogicalSection section = Assert.Single(logical.Sections);
        Assert.Same(section, logical.GetOwningSection(item));
        Assert.All(item.Lines, line => Assert.Same(section, logical.GetOwningSection(line)));

        OfficeDocumentReadResult reader = PdfReaderAdapter.ReadDocument(logical);
        OfficeDocumentBlock block = Assert.Single(reader.Blocks, static block => block.Kind == "list-item");
        Assert.Equal(item.Text, block.Text);
        Assert.DoesNotContain(reader.Blocks, static block => block.Text == "continuation tagged list line");

        string markdown = logical.ToMarkdown();
        Assert.Single(markdown.Split('\n'), static line => line.StartsWith("- ", StringComparison.Ordinal));
        Assert.Contains(item.Text, markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void LogicalProjection_KeepsSeparateVisualBulletsIndependentInsideOneHeuristicRegion() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("- First independent item", "F1", 12D, 72D, 700D, 130D),
            new PdfTextSpan("- Second independent item", "F1", 12D, 72D, 680D, 140D)
        });
        options.PageSegmentation = new WholePageRegionStage();
        options.ReadingOrder = new IdentityReadingOrderStage();

        PdfDocumentReadResult logical = Read(pdf, options);
        PdfLogicalListItem[] items = Assert.Single(logical.Pages).ListItems.ToArray();

        Assert.Equal(2, items.Length);
        Assert.Equal(new[] { "First independent item", "Second independent item" },
            items.Select(static item => item.Text));
        Assert.All(items, static item => Assert.Single(item.Lines));
        Assert.Equal(2, logical.ToMarkdown().Split('\n').Count(static line => line.StartsWith("- ", StringComparison.Ordinal)));
    }

    [Fact]
    public void LogicalProjection_KeepsHeuristicWrappedListContinuationInOneItem() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("- Wrapped list item", "F1", 12D, 72D, 700D, 110D),
            new PdfTextSpan("continuation line", "F1", 12D, 90D, 680D, 90D)
        });
        options.PageSegmentation = new WholePageRegionStage();
        options.ReadingOrder = new IdentityReadingOrderStage();

        PdfDocumentReadResult logical = Read(pdf, options);
        PdfLogicalListItem item = Assert.Single(Assert.Single(logical.Pages).ListItems);

        Assert.Equal(2, item.Lines.Count);
        Assert.Equal("Wrapped list item continuation line", item.Text);
        Assert.Single(PdfLogicalReadingOrderAnalysis.Analyze(logical.Pages[0]),
            static candidate => candidate.Kind == PdfLogicalReadingOrderKind.ListItem);
        Assert.Single(logical.ToMarkdown().Split('\n'), static line => line.StartsWith("- ", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentEnrichment_KeepsPartialTaggedRegionsAlignedWithCanonicalOrder() {
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.PageSegmentation = new WholePageRegionStage();
        options.ReadingOrder = new IdentityReadingOrderStage();
        options.SemanticClassification = new ParagraphClassificationStage();

        PdfUnderstandingPageResult page = Assert.Single(Read(CreatePartiallyTaggedRegionPdf(), options).Pages).Analysis;

        Assert.Equal(2, page.Elements.Count);
        Assert.Equal(page.Elements.Count, page.Regions.Count);
        Assert.Equal(page.Elements.Count, page.ReadingOrder.Count);
        Assert.Equal(page.Elements.Count, page.ReadingOrderEvidence.Count);
        Assert.Equal(new[] { "Tagged heading", "Untagged continuation" },
            page.Elements.Select(static element => element.Region.Text));
        Assert.Equal(PdfUnderstandingSemanticKind.Heading, page.Elements[0].Kind);
        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph, page.Elements[1].Kind);
        for (int index = 0; index < page.Elements.Count; index++) {
            Assert.Same(page.Elements[index].Region, page.Regions[index]);
            Assert.Same(page.Elements[index].Region, page.ReadingOrder[index]);
            Assert.Same(page.Elements[index].Region, page.ReadingOrderEvidence[index].Region);
            Assert.Equal(index, page.ReadingOrderEvidence[index].Index);
        }
    }

    [Fact]
    public void AdvancedPipeline_UsesCanonicalRegularFontProseTableRegions() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Assigned owner", "Current workflow" },
                new[] { "North region coordinator", "Review pending requests" },
                new[] { "South region coordinator", "Approve completed requests" }
            }, style: new PdfTableStyle {
                HeaderBold = false,
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 220D, 220D }
            })
            .ToBytes();

        PdfUnderstandingPageResult page = Assert.Single(Read(
            pdf,
            PdfUnderstandingPipelineOptions.Structured()).Pages).Analysis;

        PdfUnderstandingSemanticElement table = Assert.Single(
            page.Elements,
            static element => element.Kind == PdfUnderstandingSemanticKind.Table);
        Assert.Contains("North region coordinator", table.Region.Text, StringComparison.Ordinal);
        Assert.Contains("Approve completed requests", table.Region.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void AdvancedPipeline_PrioritizesCanonicalTablesAtPageEdges() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Item", "Helvetica-Bold", 11D, 50D, 820D, 40D),
            new PdfTextSpan("Amount", "Helvetica-Bold", 11D, 220D, 820D, 55D),
            new PdfTextSpan("Licenses", "Helvetica", 11D, 50D, 802D, 55D),
            new PdfTextSpan("42", "Helvetica", 11D, 220D, 802D, 16D)
        });

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(
            PdfUnderstandingSemanticKind.Table,
            Assert.Single(page.Elements, static element => element.Region.Text.Contains("Licenses", StringComparison.Ordinal)).Kind);
    }

    [Fact]
    public void AdvancedPipeline_TableOwnershipRequiresMeaningfulHorizontalOverlap() {
        PdfUnderstandingLine clipped = CreateUnderstandingLine("Adjacent prose", 249.5D, 339.5D, 475D);
        PdfUnderstandingLine owned = CreateUnderstandingLine("Table prose", 100D, 190D, 475D);

        Assert.False(PdfAdvancedUnderstandingStages.HasMeaningfulTableOverlap(clipped, 500D, 450D, 50D, 250D));
        Assert.True(PdfAdvancedUnderstandingStages.HasMeaningfulTableOverlap(owned, 500D, 450D, 50D, 250D));
    }

    [Fact]
    public void AdvancedReadingOrder_UsesRotatedSourceRunExtents() {
        var sourceRun = new PdfTextSpan("Vertical section label", "Helvetica", 11D, 280D, 200D, 400D, rotationDegrees: 90D);
        var word = new PdfUnderstandingWord(
            sourceRun.Text,
            280D,
            280D,
            200D,
            11D,
            90D,
            new[] { sourceRun });
        var region = new PdfUnderstandingRegion(new[] { new PdfUnderstandingLine(new[] { word }) });

        (double left, double right, double bottom, double top, _) = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(region);

        Assert.True(top - bottom >= 399D, $"Expected the rotated extent to span about 400 points, but it spanned {top - bottom:0.###}.");
        Assert.True(right - left >= 10D, $"Expected the rotated glyph thickness to span about one font size, but it spanned {right - left:0.###}.");
    }

    [Fact]
    public void AdvancedReadingOrder_LimitsSourceRunExtentToEachWordSegment() {
        var sourceRun = new PdfTextSpan("Left Right", "Helvetica", 11D, 50D, 500D, 500D);
        var leftRegion = new PdfUnderstandingRegion(new[] {
            new PdfUnderstandingLine(new[] {
                new PdfUnderstandingWord("Left", 50D, 250D, 500D, 11D, 0D, new[] { sourceRun })
            })
        });
        var rightRegion = new PdfUnderstandingRegion(new[] {
            new PdfUnderstandingLine(new[] {
                new PdfUnderstandingWord("Right", 350D, 550D, 500D, 11D, 0D, new[] { sourceRun })
            })
        });

        var leftBounds = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(leftRegion);
        var rightBounds = PdfRecursiveXyCutReadingOrderStage.GetSourceBounds(rightRegion);

        Assert.True(leftBounds.Right < rightBounds.Left, $"Expected the split word segments to retain their gutter, but bounds overlap at {leftBounds.Right:0.###} and {rightBounds.Left:0.###}.");
    }

    [Fact]
    public void AdvancedReadingOrder_IsolatesCloseSpanningHeadingBeforeColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(sourcePage, 1, new PdfTextLayoutOptions(), 10000, 10000);
        PdfUnderstandingRegion heading = new(new[] { CreateUnderstandingLine("Spanning heading", 50D, 500D, 710D) });
        PdfUnderstandingRegion leftTop = new(new[] { CreateUnderstandingLine("Left top", 50D, 160D, 700D) });
        PdfUnderstandingRegion leftBottom = new(new[] { CreateUnderstandingLine("Left bottom", 50D, 160D, 650D) });
        PdfUnderstandingRegion rightTop = new(new[] { CreateUnderstandingLine("Right top", 320D, 430D, 700D) });
        PdfUnderstandingRegion rightBottom = new(new[] { CreateUnderstandingLine("Right bottom", 320D, 430D, 650D) });

        IReadOnlyList<PdfUnderstandingRegion> ordered = new PdfRecursiveXyCutReadingOrderStage().Order(
            context,
            new[] { heading, leftTop, rightTop, leftBottom, rightBottom });

        Assert.Equal(
            new[] { "Spanning heading", "Left top", "Left bottom", "Right top", "Right bottom" },
            ordered.Select(static region => region.Text));
    }

    [Fact]
    public void AdvancedReadingOrder_IsolatesCloseTrailingSpanningRegionAfterColumns() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage sourcePage = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(sourcePage, 1, new PdfTextLayoutOptions(), 10000, 10000);
        PdfUnderstandingRegion leftTop = new(new[] { CreateUnderstandingLine("Left top", 50D, 160D, 700D) });
        PdfUnderstandingRegion leftBottom = new(new[] { CreateUnderstandingLine("Left bottom", 50D, 160D, 650D) });
        PdfUnderstandingRegion rightTop = new(new[] { CreateUnderstandingLine("Right top", 320D, 430D, 700D) });
        PdfUnderstandingRegion rightBottom = new(new[] { CreateUnderstandingLine("Right bottom", 320D, 430D, 650D) });
        PdfUnderstandingRegion trailing = new(new[] { CreateUnderstandingLine("Trailing table", 50D, 500D, 640D) });

        IReadOnlyList<PdfUnderstandingRegion> ordered = new PdfRecursiveXyCutReadingOrderStage().Order(
            context,
            new[] { leftTop, rightTop, leftBottom, rightBottom, trailing });

        Assert.Equal(
            new[] { "Left top", "Left bottom", "Right top", "Right bottom", "Trailing table" },
            ordered.Select(static region => region.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pipeline_DoesNotClassifyDecimalAmountsAsListItems(bool advanced) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var glyphs = new FixedGlyphStage(new[] {
            new PdfTextSpan("1037.25", "F1", 11, 50, 500, 45),
            new PdfTextSpan("1. Actual numbered item", "F1", 11, 50, 430, 120),
            new PdfTextSpan("-42", "F1", 11, 50, 360, 24)
        });
        PdfUnderstandingPipelineOptions options = advanced
            ? PdfUnderstandingPipelineOptions.Structured()
            : new PdfUnderstandingPipelineOptions();
        options.GlyphDecoding = glyphs;

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements, element => element.Region.Text == "1037.25").Kind);
        Assert.Equal(PdfUnderstandingSemanticKind.ListItem,
            Assert.Single(page.Elements, element => element.Region.Text == "1. Actual numbered item").Kind);
        Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements, element => element.Region.Text == "-42").Kind);
    }

    [Theory]
    [InlineData(false, "1.2. Nested numbered item")]
    [InlineData(true, "1.2. Nested numbered item")]
    [InlineData(false, "2.3.1)Deep numbered item")]
    [InlineData(true, "2.3.1)Deep numbered item")]
    [InlineData(false, "3.Compact numbered item")]
    [InlineData(true, "3.Compact numbered item")]
    [InlineData(false, "(a)Compact parenthesized item")]
    [InlineData(true, "(a)Compact parenthesized item")]
    [InlineData(false, "(1)Compact numeric parenthesized item")]
    [InlineData(true, "(1)Compact numeric parenthesized item")]
    [InlineData(false, "-Compact ASCII bullet")]
    [InlineData(true, "-Compact ASCII bullet")]
    [InlineData(false, "*Compact ASCII bullet")]
    [InlineData(true, "*Compact ASCII bullet")]
    public void Pipeline_ClassifiesHierarchicalAndCompactListItems(bool advanced, string text) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        PdfUnderstandingPipelineOptions options = advanced
            ? PdfUnderstandingPipelineOptions.Structured()
            : new PdfUnderstandingPipelineOptions();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan(text, "F1", 11, 50, 500, 180)
        });

        PdfUnderstandingPageResult page = Assert.Single(Read(pdf, options).Pages).Analysis;

        Assert.Equal(PdfUnderstandingSemanticKind.ListItem,
            Assert.Single(page.Elements, element => element.Region.Text == text).Kind);
    }

    [Theory]
    [InlineData("-$42 total")]
    [InlineData("-€42 total")]
    [InlineData("-£ 42 total")]
    public void SharedListParser_DoesNotClassifySignedCurrencyAsCompactBullets(string text) {
        Assert.False(ContentStructureExtractor.IsListItemText(text));
    }

    [Theory]
    [InlineData("-.5 variance")]
    [InlineData("-,25 margin")]
    [InlineData("-$.5 variance")]
    public void SharedListParser_DoesNotClassifyLeadingDecimalValuesAsCompactBullets(string text) {
        Assert.False(ContentStructureExtractor.IsListItemText(text));
    }

    [Theory]
    [InlineData("--output path")]
    [InlineData("**bold**")]
    [InlineData("-*literal")]
    public void SharedListParser_DoesNotClassifyRepeatedPunctuationAsCompactBullets(string text) {
        Assert.False(ContentStructureExtractor.IsListItemText(text));
    }

    private sealed class ReverseReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions) => regions.Reverse().ToArray();
    }

    private static PdfDocumentReadResult Read(
        byte[] pdf,
        PdfUnderstandingPipelineOptions pipeline,
        PdfReadProfile profile = PdfReadProfile.Structured,
        PdfPageSelection? selection = null,
        PdfTextLayoutOptions? layoutOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfDocument.Load(pdf).Read(new PdfReadOptions {
            Profile = profile,
            PageSelection = selection,
            LayoutOptions = layoutOptions ?? new PdfTextLayoutOptions(),
            Pipeline = pipeline
        }, cancellationToken);
    }

    private static byte[] CreateNumberedOutlinePdf() {
        string content = "BT /F1 24 Tf 72 700 Td (Section 1) Tj ET\n" +
            "BT /F1 14 Tf 72 650 Td (Section 2) Tj ET\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /Outlines 6 0 R /PageMode /UseOutlines >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, content),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /Outlines /First 7 0 R /Last 7 0 R /Count 1 >>",
            "<< /Title (Section 2) /Parent 6 0 R /Dest [3 0 R /Fit] >>");
    }

    private static byte[] CreateScopedTaggedMcidPdf() {
        string pageContent = "/H1 << /MCID 0 >> BDC\n" +
            "BT /F1 24 Tf 72 700 Td (Page heading) Tj ET\nEMC\n/Fx1 Do\n";
        string formContent = "/P << /MCID 0 >> BDC\n" +
            "BT /F1 12 Tf 72 500 Td (Form paragraph) Tj ET\nEMC\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 7 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 " +
                "/Resources << /Font << /F1 5 0 R >> /XObject << /Fx1 6 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, pageContent),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            BuildStreamBody("/Type /XObject /Subtype /Form /BBox [0 0 612 792] /StructParent 1 " +
                "/Resources << /Font << /F1 5 0 R >> >>", formContent),
            "<< /Type /StructTreeRoot /K [8 0 R 9 0 R] /ParentTree 10 0 R /ParentTreeNextKey 2 >>",
            "<< /Type /StructElem /S /H1 /P 7 0 R /Pg 3 0 R /K 0 >>",
            "<< /Type /StructElem /S /P /P 7 0 R /Pg 3 0 R " +
                "/K << /Type /MCR /Pg 3 0 R /Stm 6 0 R /MCID 0 >> >>",
            "<< /Nums [0 [8 0 R] 1 [9 0 R]] >>");
    }

    private static byte[] CreateMultiLineTaggedListItemPdf() {
        string content = "/H1 << /MCID 0 >> BDC\n" +
            "BT /F1 18 Tf 72 740 Td (Tagged list section) Tj ET\nEMC\n" +
            "/Span << /MCID 1 >> BDC\n" +
            "BT /F1 12 Tf 72 700 Td (First tagged list line) Tj ET\nEMC\n" +
            "/Span << /MCID 2 >> BDC\n" +
            "BT /F1 12 Tf 90 680 Td (continuation tagged list line) Tj ET\nEMC\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 " +
                "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, content),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /StructTreeRoot /K [7 0 R 8 0 R] /ParentTree 9 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /H1 /P 6 0 R /Pg 3 0 R /K 0 >>",
            "<< /Type /StructElem /S /LI /P 6 0 R /Pg 3 0 R /K [1 2] >>",
            "<< /Nums [0 [7 0 R 8 0 R 8 0 R]] >>");
    }

    private static byte[] CreatePartiallyTaggedRegionPdf() {
        string content = "/H2 << /MCID 0 >> BDC\n" +
            "BT /F1 18 Tf 72 700 Td (Tagged heading) Tj ET\nEMC\n" +
            "BT /F1 12 Tf 72 680 Td (Untagged continuation) Tj ET\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 " +
                "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, content),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 8 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /H2 /P 6 0 R /Pg 3 0 R /K 0 >>",
            "<< /Nums [0 [7 0 R]] >>");
    }

    private static byte[] CreateInvalidTaggedHeadingPdf() {
        string content = "/H0 << /MCID 0 >> BDC\n" +
            "BT /F1 12 Tf 72 700 Td (Malformed tagged heading) Tj ET\nEMC\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 " +
                "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, content),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 8 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /H0 /P 6 0 R /Pg 3 0 R /K 0 >>",
            "<< /Nums [0 [7 0 R]] >>");
    }

    private static byte[] CreateNestedTaggedHeadingPdf() {
        string content = "/Span << /MCID 0 >> BDC\n" +
            "BT /F1 18 Tf 72 700 Td (Nested tagged heading) Tj ET\nEMC\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 " +
                "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, content),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 9 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /H1 /P 6 0 R /Pg 3 0 R /K [8 0 R] >>",
            "<< /Type /StructElem /S /Span /P 7 0 R /Pg 3 0 R /K 0 >>",
            "<< /Nums [0 [8 0 R]] >>");
    }

    private static byte[] CreateCyclicTaggedStructurePdf() {
        string content = "/Span << /MCID 0 >> BDC\n" +
            "BT /F1 12 Tf 72 700 Td (Cyclic tagged content) Tj ET\nEMC\n";
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 " +
                "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, content),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 9 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /Span /P 8 0 R /Pg 3 0 R /K 0 >>",
            "<< /Type /StructElem /S /Div /P 7 0 R /Pg 3 0 R /K [] >>",
            "<< /Nums [0 [7 0 R]] >>");
    }

    private static string BuildStreamBody(string dictionaryEntries, string content) {
        int length = System.Text.Encoding.ASCII.GetByteCount(content);
        return "<< " + dictionaryEntries + " /Length " +
            length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n" + content + "endstream";
    }

    private static byte[] BuildClassicPdf(params string[] objectBodies) {
        using var output = new MemoryStream();
        var offsets = new long[objectBodies.Length + 1];
        WriteAscii(output, "%PDF-1.7\n");
        for (int index = 0; index < objectBodies.Length; index++) {
            int objectNumber = index + 1;
            offsets[objectNumber] = output.Position;
            WriteAscii(output, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " 0 obj\n" + objectBodies[index] + "\nendobj\n");
        }

        long xrefOffset = output.Position;
        WriteAscii(output, "xref\n0 " + (objectBodies.Length + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "\n0000000000 65535 f \n");
        for (int objectNumber = 1; objectNumber <= objectBodies.Length; objectNumber++) {
            WriteAscii(output, offsets[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) +
                " 00000 n \n");
        }
        WriteAscii(output, "trailer\n<< /Size " + (objectBodies.Length + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /Root 1 0 R >>\nstartxref\n" + xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static PdfUnderstandingPipelineOptions CreatePassThroughPipeline(IPdfGlyphDecodingStage glyphStage, long maxWorkUnits) =>
        new PdfUnderstandingPipelineOptions {
            GlyphDecoding = glyphStage,
            WordGrouping = new SingleWordGroupingStage(),
            LineGrouping = new SingleLineGroupingStage(),
            PageSegmentation = new SingleRegionStage(),
            ReadingOrder = new IdentityReadingOrderStage(),
            SemanticClassification = new ParagraphClassificationStage(),
            MaxWorkUnitsPerPage = maxWorkUnits
        };

    private static void AssertTaggedHeading(PdfUnderstandingPageResult page, string text, int level) {
        Assert.Contains(page.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Heading &&
            element.Level == level &&
            element.Region.Text.Contains(text, StringComparison.Ordinal) &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.tagged-pdf-role"));
    }

    private static PdfUnderstandingLine CreateGappedUnderstandingLine(string left, string right, double baselineY) {
        var leftRun = new PdfTextSpan(left, "Helvetica", 11D, 50D, baselineY, 60D);
        var rightRun = new PdfTextSpan(right, "Helvetica", 11D, 250D, baselineY, 60D);
        return new PdfUnderstandingLine(new[] {
            new PdfUnderstandingWord(left, 50D, 110D, baselineY, 11D, 0D, new[] { leftRun }),
            new PdfUnderstandingWord(right, 250D, 310D, baselineY, 11D, 0D, new[] { rightRun })
        });
    }

    private static byte[] CreateEmptyTaggedStructurePdf() {
        return BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 5 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /StructTreeRoot /K [6 0 R] >>",
            "<< /Type /StructElem /S /Div /P 5 0 R /Pg 3 0 R /K [] >>");
    }

    private static PdfUnderstandingLine CreateUnderstandingLine(
        string text,
        double xStart,
        double xEnd,
        double baselineY) {
        var run = new PdfTextSpan(text, "Helvetica", 11D, xStart, baselineY, xEnd - xStart);
        var word = new PdfUnderstandingWord(text, xStart, xEnd, baselineY, 11D, 0D, new[] { run });
        return new PdfUnderstandingLine(new[] { word });
    }

    private static PdfUnderstandingWord CreateUnderstandingWord(string text, double xStart, double baselineY) {
        var run = new PdfTextSpan(text, "Helvetica", 11D, xStart, baselineY, 40D);
        return new PdfUnderstandingWord(text, xStart, xStart + 40D, baselineY, 11D, 0D, new[] { run });
    }

    private static PdfUnderstandingRegion CreateUnderstandingRegion(string text, double xStart, double baselineY, double fontSize) {
        var run = new PdfTextSpan(text, "Helvetica", fontSize, xStart, baselineY, 40D);
        var word = new PdfUnderstandingWord(text, xStart, xStart + 40D, baselineY, fontSize, 0D, new[] { run });
        return new PdfUnderstandingRegion(new[] { new PdfUnderstandingLine(new[] { word }) });
    }

    private sealed class FixedGlyphStage : IPdfGlyphDecodingStage {
        private readonly IReadOnlyList<PdfTextSpan> _spans;
        internal FixedGlyphStage(IReadOnlyList<PdfTextSpan> spans) { _spans = spans; }
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) => _spans;
    }

    private sealed class BudgetChargingGlyphStage : IPdfGlyphDecodingStage {
        private readonly long _workUnits;
        internal BudgetChargingGlyphStage(long workUnits) { _workUnits = workUnits; }
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) {
            context.ConsumeWork(_workUnits);
            return new[] { new PdfTextSpan("word", "F1", 12D, 10D, 10D, 24D) };
        }
    }

    private sealed class CancellingGlyphStage : IPdfGlyphDecodingStage {
        private readonly CancellationTokenSource _cancellation;
        private readonly int _cancelAfterWorkUnit;
        internal CancellingGlyphStage(CancellationTokenSource cancellation, int cancelAfterWorkUnit) {
            _cancellation = cancellation;
            _cancelAfterWorkUnit = cancelAfterWorkUnit;
        }
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) {
            for (int index = 0; index < 20; index++) {
                if (index == _cancelAfterWorkUnit) _cancellation.Cancel();
                context.ConsumeWork();
            }
            return Array.Empty<PdfTextSpan>();
        }
    }

    private sealed class SingleWordGroupingStage : IPdfWordGroupingStage {
        public IReadOnlyList<PdfUnderstandingWord> GroupWords(PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
            PdfTextSpan run = Assert.Single(runs);
            return new[] { new PdfUnderstandingWord(run.Text, run.X, run.X + run.Advance, run.Y, run.FontSize, run.RotationDegrees, new[] { run }) };
        }
    }

    private sealed class SingleLineGroupingStage : IPdfLineGroupingStage {
        public IReadOnlyList<PdfUnderstandingLine> GroupLines(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingWord> words) =>
            new[] { new PdfUnderstandingLine(new[] { Assert.Single(words) }) };
    }

    private sealed class FixedLineGroupingStage : IPdfLineGroupingStage {
        private readonly IReadOnlyList<PdfUnderstandingLine> _lines;
        internal FixedLineGroupingStage(IReadOnlyList<PdfUnderstandingLine> lines) { _lines = lines; }
        public IReadOnlyList<PdfUnderstandingLine> GroupLines(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingWord> words) => _lines;
    }

    private sealed class CombinedLineGroupingStage : IPdfLineGroupingStage {
        public IReadOnlyList<PdfUnderstandingLine> GroupLines(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingWord> words) => words.Count == 0
                ? Array.Empty<PdfUnderstandingLine>()
                : new[] { new PdfUnderstandingLine(words) };
    }

    private sealed class SingleRegionStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) =>
            new[] { new PdfUnderstandingRegion(new[] { Assert.Single(lines) }) };
    }

    private sealed class WholePageRegionStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) =>
            lines.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : new[] { new PdfUnderstandingRegion(lines) };
    }

    private sealed class HeadingThenRemainingRegionStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) {
            if (lines.Count <= 1) return lines.Count == 0
                ? Array.Empty<PdfUnderstandingRegion>()
                : new[] { new PdfUnderstandingRegion(lines) };
            return new[] {
                new PdfUnderstandingRegion(new[] { lines[0] }),
                new PdfUnderstandingRegion(lines.Skip(1).ToArray())
            };
        }
    }

    private sealed class IdentityReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions) => regions;
    }

    private sealed class ParagraphClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) =>
            orderedRegions.Select(static region => new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Paragraph)).ToArray();
    }

    private sealed class HeadingClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) =>
            orderedRegions.Select(static region => new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Heading, level: 1)).ToArray();
    }

    private sealed class HeadingLevelTwoClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) =>
            orderedRegions.Select(static region => region.Lines.Max(static line => line.FontSize) >= 15D
                ? new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Heading, level: 2)
                : new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Paragraph)).ToArray();
    }

    private sealed class TableSectionClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingRegion> orderedRegions) =>
            orderedRegions.Select(static region => region.Text.StartsWith("Table section", StringComparison.Ordinal)
                ? new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Heading, level: 1)
                : region.Evidence.Any(static evidence => evidence.Code == "region.canonical-table")
                    ? new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Table)
                    : new PdfUnderstandingSemanticElement(region, PdfUnderstandingSemanticKind.Paragraph)).ToArray();
    }

    private sealed class UnmarkedListClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) =>
            orderedRegions.Select(static region => new PdfUnderstandingSemanticElement(
                region,
                PdfUnderstandingSemanticKind.ListItem,
                0.95D,
                new[] { new PdfInferenceEvidence("semantic.test-list-item", "The test classifier supplied authoritative list semantics.", 1D) },
                level: 2)).ToArray();
    }
}
