using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSemanticAuthoringTests {
    [Fact]
    public void SemanticGroupsCreateNestedTaggedStructureWithoutChangingFlowText() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers)
            .Content(content => content.Semantic(PdfSemanticRole.Article, article => article
                .Semantic(PdfSemanticRole.Division, division => division
                    .H1("Accessible report")
                    .Text("Operational summary")))))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo article = Assert.Single(tagged.StructureElements, element => element.StructureType == "Art");
        PdfStructureElementInfo division = Assert.Single(tagged.StructureElements, element => element.StructureType == "Div");
        PdfStructureElementInfo heading = Assert.Single(tagged.StructureElements, element => element.StructureType == "H1");
        PdfStructureElementInfo paragraph = Assert.Single(tagged.StructureElements, element => element.StructureType == "P");

        Assert.Contains(division.ObjectNumber, article.ChildElementObjectNumbers);
        Assert.Contains(heading.ObjectNumber, division.ChildElementObjectNumbers);
        Assert.Contains(paragraph.ObjectNumber, division.ChildElementObjectNumbers);
        Assert.Contains("Accessible report", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
        Assert.Contains("Operational summary", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void FigureSemanticsRequireAndPreserveAlternateText() {
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(document => document.Content(content =>
            content.Semantic(PdfSemanticRole.Figure, figure => figure.Text("Chart")))));

        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers)
            .Content(content => content.Semantic(
                PdfSemanticRole.Figure,
                figure => figure.Text("Quarterly trend"),
                alternativeText: "Line chart showing a rising quarterly trend")))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo figure = Assert.Single(tagged.StructureElements, element => element.StructureType == "Figure");

        Assert.Equal("Line chart showing a rising quarterly trend", figure.AlternateText);
        Assert.Contains(figure.ChildElementObjectNumbers.Single(), tagged.StructureElements.Select(element => element.ObjectNumber));
    }

    [Fact]
    public void DocumentSectionsCreateSectionStructureAlongsideOutlineDestinations() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers)
            .Content(content => content.Section("Results", section => section.Text("Complete"))))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, element => element.StructureType == "Sect");
        PdfStructureElementInfo heading = Assert.Single(tagged.StructureElements, element => element.StructureType == "H1");
        PdfStructureElementInfo paragraph = Assert.Single(tagged.StructureElements, element => element.StructureType == "P");

        Assert.Contains(heading.ObjectNumber, section.ChildElementObjectNumbers);
        Assert.Contains(paragraph.ObjectNumber, section.ChildElementObjectNumbers);
    }

    [Fact]
    public void DecoratedElementCombinesLayoutAndSemanticHierarchy() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                options.CompressContentStreams = false;
            })
            .Content(content => content.Element(element => element
                .Semantic(PdfSemanticRole.BlockQuote)
                .Background(new PdfColor(0.9D, 0.95D, 1D))
                .Border(PdfColor.FromRgb(90, 100, 115))
                .Padding(8)
                .Content(quote => quote.Text("Measured twice, rendered once.")))))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo quote = Assert.Single(tagged.StructureElements, element => element.StructureType == "BlockQuote");
        PdfStructureElementInfo paragraph = Assert.Single(tagged.StructureElements, element => element.StructureType == "P");
        string raw = PdfEncoding.Latin1GetString(bytes);

        Assert.Contains(paragraph.ObjectNumber, quote.ChildElementObjectNumbers);
        Assert.Contains("0.9 0.95 1 rg", raw, StringComparison.Ordinal);
        Assert.Contains("Measured twice, rendered once.", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DecoratedElementAcceptsReusableComponentsThroughTheSameContentReceiver() {
        byte[] bytes = PdfDocument.Create(document => document.Content(content => content
            .Element(element => element
                .Padding(6)
                .Content(nested => nested.Component(new TestComponent())))))
            .ToBytes();

        Assert.Contains("Reusable element content", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DecoratedElementCopiesReusableStylesBeforeCompositionContinues() {
        var style = new PdfPanelStyle {
            Background = new PdfColor(0.1D, 0.2D, 0.3D),
            PaddingX = 5D,
            PaddingY = 5D
        };

        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => options.CompressContentStreams = false)
            .Content(content => content.Element(element => element
                .Style(style)
                .Content(nested => nested.Text("Reusable style")))))
            .ToBytes();
        style.Background = PdfColor.White;

        Assert.Contains("0.1 0.2 0.3 rg", PdfEncoding.Latin1GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void DecoratedElementAcceptsSizedRowsThroughTheNormalFlowEngine() {
        byte[] bytes = PdfDocument.Create(document => document
            .Content(content => content.Element(element => element
                .Padding(6)
                .Content(nested => nested.Row(row => row
                    .FixedColumn(48, column => column.Text("Fixed"))
                    .AutoColumn(column => column.Text("Automatic"))
                    .RelativeColumn(column => column.Text("Remaining")))))))
            .ToBytes();

        string text = PdfDocument.Load(bytes).Read().Text;
        Assert.Contains("Fixed", text, StringComparison.Ordinal);
        Assert.Contains("Automatic", text, StringComparison.Ordinal);
        Assert.Contains("Remaining", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DecoratedElementContinuesRowsAndDecorationAcrossPages() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.PageWidth = 300;
                options.PageHeight = 220;
                options.Margins = PageMargins.Uniform(24);
                options.CompressContentStreams = false;
            })
            .Content(content => content.Element(element => element
                .Background(new PdfColor(0.9D, 0.95D, 1D))
                .Border(PdfColor.FromRgb(90, 100, 115))
                .Padding(6)
                .Content(nested => nested.Row(row => row
                    .RelativeColumn(column => {
                        for (int index = 0; index < 24; index++) {
                            column.Text("Left " + index);
                        }
                    })
                    .RelativeColumn(column => {
                        for (int index = 0; index < 24; index++) {
                            column.Text("Right " + index);
                        }
                    }))))))
            .ToBytes();

        int pageCount = PdfInspector.Inspect(bytes).PageCount;
        string raw = PdfEncoding.Latin1GetString(bytes);
        string text = PdfDocument.Load(bytes).Read().Text;

        Assert.True(pageCount > 1);
        Assert.True(System.Text.RegularExpressions.Regex.Matches(raw, "0.9 0.95 1 rg").Count >= pageCount);
        Assert.Contains("Left 23", text, StringComparison.Ordinal);
        Assert.Contains("Right 23", text, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepTogetherMeasuresContentInsideSemanticGroups() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfDocument.Create(document => document
            .Settings(options => {
                options.PageWidth = 300;
                options.PageHeight = 180;
                options.Margins = PageMargins.Uniform(24);
            })
            .Content(content => content.Element(element => element
                .Padding(6)
                .KeepTogether()
                .Content(nested => nested.Semantic(PdfSemanticRole.Article, article => {
                    for (int index = 0; index < 20; index++) {
                        article.Text("A complete line of measured semantic content " + index);
                    }
                })))))
            .ToBytes());

        Assert.Contains("exceeds the available page content height", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FlowKeepTogetherMeasuresCompleteSemanticGroups() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfDocument.Create(document => document
            .Settings(options => {
                options.PageWidth = 300;
                options.PageHeight = 180;
                options.Margins = PageMargins.Uniform(24);
            })
            .Content(content => content.Flow(flow => flow.Semantic(PdfSemanticRole.Article, article => {
                for (int index = 0; index < 20; index++) {
                    article.Text("A complete line of grouped flow content " + index);
                }
            }), new PdfFlowOptions { KeepTogether = true })))
            .ToBytes());

        Assert.Contains("exceeds the available full-page content height", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepTogetherMeasuresNestedSpacingWithTheRenderedCursorRules() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfDocument.Create(document => document
            .Settings(options => {
                options.PageWidth = 300;
                options.PageHeight = 180;
                options.Margins = PageMargins.Uniform(30);
            })
            .Content(content => content.Element(element => element
                .KeepTogether()
                .Content(nested => nested
                    .PanelParagraph(paragraph => paragraph.Text("First"), new PdfPanelStyle { PaddingY = 2, SpacingAfter = 0 })
                    .PanelParagraph(paragraph => paragraph.Text("Second"), new PdfPanelStyle { PaddingY = 2, SpacingBefore = 90, SpacingAfter = 0 })))))
            .ToBytes());

        Assert.Contains("exceeds the available page content height", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepTogetherDoesNotApplyTopSpacingThatRenderingSuppresses() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.PageWidth = 300;
                options.PageHeight = 180;
                options.Margins = PageMargins.Uniform(30);
            })
            .Content(content => content.Element(outer => outer
                .KeepTogether()
                .Content(nested => nested.Element(inner => inner
                    .Spacing(before: 110, after: 0)
                    .Content(body => body.Text("Top spacing is suppressed")))))))
            .ToBytes();

        Assert.Equal(1, PdfInspector.Inspect(bytes).PageCount);
        Assert.Contains("Top spacing is suppressed", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ElementKeepWithNextParticipatesInTheCompleteKeepChain() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.PageWidth = 260;
                options.PageHeight = 170;
                options.Margins = PageMargins.Uniform(30);
                options.DefaultFontSize = 10;
            })
            .Content(content => content
                .Paragraph(paragraph => paragraph.Text("IntroMarker"), style: new PdfParagraphStyle { SpacingAfter = 48 })
                .H3("ElementChainHeading")
                .Element(element => element
                    .Padding(4)
                    .Spacing(after: 0)
                    .KeepWithNext()
                    .Content(nested => nested.Text("ElementChainBody")))
                .Text("ElementChainFollowing")))
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("ElementChainHeading", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("ElementChainHeading", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("ElementChainBody", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("ElementChainFollowing", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ElementKeepWithNextRejectsUnmeasurableContent() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfDocument.Create(document => document
            .Content(content => content
                .Element(element => element
                    .KeepWithNext()
                    .Content(nested => nested.Deferred(_ => deferred => deferred.Text("Page-aware element"))))
                .Text("Following content")))
            .ToBytes());

        Assert.Contains("height can be determined before rendering", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SemanticGroupRemainsOneLogicalStructureElementAcrossPages() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                options.PageWidth = 300;
                options.PageHeight = 180;
                options.Margins = PageMargins.Uniform(24);
            })
            .Content(content => content.Semantic(PdfSemanticRole.Article, article => {
                for (int index = 0; index < 24; index++) {
                    article.Text("Cross-page semantic paragraph " + index);
                }
            })))
            .ToBytes();

        Assert.True(PdfInspector.Inspect(bytes).PageCount > 1);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo article = Assert.Single(tagged.StructureElements, element => element.StructureType == "Art");
        PdfStructureElementInfo[] paragraphs = tagged.StructureElements.Where(element => element.StructureType == "P").ToArray();

        Assert.Equal(24, paragraphs.Length);
        Assert.All(paragraphs, paragraph => Assert.Contains(paragraph.ObjectNumber, article.ChildElementObjectNumbers));
    }

    [Fact]
    public void CrossPageFigureSemanticsEmitAlternateTextOnce() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                options.PageWidth = 300;
                options.PageHeight = 180;
                options.Margins = PageMargins.Uniform(24);
            })
            .Content(content => content.Semantic(PdfSemanticRole.Figure, figure => {
                for (int index = 0; index < 24; index++) {
                    figure.Text("Figure narrative " + index);
                }
            }, alternativeText: "One cross-page figure")))
            .ToBytes();

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo figure = Assert.Single(tagged.StructureElements, element => element.StructureType == "Figure");
        string source = PdfEncoding.Latin1GetString(bytes);
        System.Text.RegularExpressions.Match figureObject = System.Text.RegularExpressions.Regex.Match(
            source,
            "(?ms)^" + figure.ObjectNumber + " 0 obj\\s*(?<dictionary><<.*?>>)\\s*endobj$");

        Assert.Equal("One cross-page figure", figure.AlternateText);
        Assert.Equal(24, figure.ChildElementObjectNumbers.Count);
        Assert.True(figureObject.Success);
        Assert.DoesNotContain("/Pg", figureObject.Groups["dictionary"].Value, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepTogetherRejectsDynamicContentInsteadOfMakingAFalsePaginationPromise() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfDocument.Create(document => document
            .Content(content => content.Element(element => element
                .KeepTogether()
                .Content(nested => nested.Deferred(_ => deferred => deferred.Text("Page-aware content"))))))
            .ToBytes());

        Assert.Contains("height can be determined before rendering", exception.Message, StringComparison.Ordinal);
    }

    private sealed class TestComponent : IPdfComponent {
        public void Compose(PdfContentBuilder content) {
            content.Text("Reusable element content");
        }
    }
}
