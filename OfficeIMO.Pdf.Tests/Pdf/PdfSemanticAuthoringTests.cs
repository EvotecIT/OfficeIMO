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
                    .AddParagraph("Operational summary")))))
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
            content.Semantic(PdfSemanticRole.Figure, figure => figure.AddParagraph("Chart")))));

        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers)
            .Content(content => content.Semantic(
                PdfSemanticRole.Figure,
                figure => figure.AddParagraph("Quarterly trend"),
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
            .Content(content => content.Section("Results", section => section.AddParagraph("Complete"))))
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
                .Background(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.9D, 0.95D, 1D))
                .Border(global::OfficeIMO.OfficeColor.FromRgb(90, 100, 115))
                .Padding(global::OfficeIMO.OfficeLength.Points(8))
                .Content(quote => quote.AddParagraph("Measured twice, rendered once.")))))
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
                .Padding(global::OfficeIMO.OfficeLength.Points(6))
                .Content(nested => nested.Component(new TestComponent())))))
            .ToBytes();

        Assert.Contains("Reusable element content", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DecoratedElementCopiesReusableStylesBeforeCompositionContinues() {
        var style = new PdfPanelStyle {
            Background = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.1D, 0.2D, 0.3D),
            PaddingX = global::OfficeIMO.OfficeLength.Points(5D),
            PaddingY = global::OfficeIMO.OfficeLength.Points(5D)
        };

        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => options.CompressContentStreams = false)
            .Content(content => content.Element(element => element
                .Style(style)
                .Content(nested => nested.AddParagraph("Reusable style")))))
            .ToBytes();
        style.Background = global::OfficeIMO.OfficeColor.White;

        Assert.Contains("0.1 0.2 0.3 rg", PdfEncoding.Latin1GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void DecoratedElementAcceptsSizedRowsThroughTheNormalFlowEngine() {
        byte[] bytes = PdfDocument.Create(document => document
            .Content(content => content.Element(element => element
                .Padding(global::OfficeIMO.OfficeLength.Points(6))
                .Content(nested => nested.Row(row => row
                    .FixedColumn(global::OfficeIMO.OfficeLength.Points(48), column => column.AddParagraph("Fixed"))
                    .AutoColumn(column => column.AddParagraph("Automatic"))
                    .RelativeColumn(column => column.AddParagraph("Remaining")))))))
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
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(220));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)));
                options.CompressContentStreams = false;
            })
            .Content(content => content.Element(element => element
                .Background(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.9D, 0.95D, 1D))
                .Border(global::OfficeIMO.OfficeColor.FromRgb(90, 100, 115))
                .Padding(global::OfficeIMO.OfficeLength.Points(6))
                .Content(nested => nested.Row(row => row
                    .RelativeColumn(column => {
                        for (int index = 0; index < 24; index++) {
                            column.AddParagraph("Left " + index);
                        }
                    })
                    .RelativeColumn(column => {
                        for (int index = 0; index < 24; index++) {
                            column.AddParagraph("Right " + index);
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
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)));
            })
            .Content(content => content.Element(element => element
                .Padding(global::OfficeIMO.OfficeLength.Points(6))
                .KeepTogether()
                .Content(nested => nested.Semantic(PdfSemanticRole.Article, article => {
                    for (int index = 0; index < 20; index++) {
                        article.AddParagraph("A complete line of measured semantic content " + index);
                    }
                })))))
            .ToBytes());

        Assert.Contains("exceeds the available page content height", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FlowKeepTogetherMeasuresCompleteSemanticGroups() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfDocument.Create(document => document
            .Settings(options => {
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)));
            })
            .Content(content => content.Flow(flow => flow.Semantic(PdfSemanticRole.Article, article => {
                for (int index = 0; index < 20; index++) {
                    article.AddParagraph("A complete line of grouped flow content " + index);
                }
            }), new PdfFlowOptions { KeepTogether = true })))
            .ToBytes());

        Assert.Contains("exceeds the available full-page content height", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepTogetherMeasuresNestedSpacingWithTheRenderedCursorRules() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfDocument.Create(document => document
            .Settings(options => {
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)));
            })
            .Content(content => content.Element(element => element
                .KeepTogether()
                .Content(nested => nested
                    .AddPanelParagraph(paragraph => paragraph.AddText("First"), new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(2), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                    .AddPanelParagraph(paragraph => paragraph.AddText("Second"), new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(2), SpacingBefore = global::OfficeIMO.OfficeLength.Points(90), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })))))
            .ToBytes());

        Assert.Contains("exceeds the available page content height", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void KeepTogetherDoesNotApplyTopSpacingThatRenderingSuppresses() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)));
            })
            .Content(content => content.Element(outer => outer
                .KeepTogether()
                .Content(nested => nested.Element(inner => inner
                    .Spacing(before: global::OfficeIMO.OfficeLength.Points(110), after: global::OfficeIMO.OfficeLength.Points(0))
                    .Content(body => body.AddParagraph("Top spacing is suppressed")))))))
            .ToBytes();

        Assert.Equal(1, PdfInspector.Inspect(bytes).PageCount);
        Assert.Contains("Top spacing is suppressed", PdfDocument.Load(bytes).Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ElementKeepWithNextParticipatesInTheCompleteKeepChain() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(260), global::OfficeIMO.OfficeLength.Points(170));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)));
                options.DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10);
            })
            .Content(content => content
                .AddParagraph(paragraph => paragraph.AddText("IntroMarker"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(48) })
                .H3("ElementChainHeading")
                .Element(element => element
                    .Padding(global::OfficeIMO.OfficeLength.Points(4))
                    .Spacing(after: global::OfficeIMO.OfficeLength.Points(0))
                    .KeepWithNext()
                    .Content(nested => nested.AddParagraph("ElementChainBody")))
                .AddParagraph("ElementChainFollowing")))
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
                    .Content(nested => nested.Deferred(_ => deferred => deferred.AddParagraph("Page-aware element"))))
                .AddParagraph("Following content")))
            .ToBytes());

        Assert.Contains("height can be determined before rendering", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SemanticGroupRemainsOneLogicalStructureElementAcrossPages() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)));
            })
            .Content(content => content.Semantic(PdfSemanticRole.Article, article => {
                for (int index = 0; index < 24; index++) {
                    article.AddParagraph("Cross-page semantic paragraph " + index);
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
    public void SemanticAncestorOfCrossPageHeadingHasNoPageReference() {
        string heading = string.Join(" ", Enumerable.Repeat("Continued heading", 90));
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)));
            })
            .Content(content => content.Semantic(PdfSemanticRole.Article, article => article.H1(heading))))
            .ToBytes();

        Assert.True(PdfInspector.Inspect(bytes).PageCount > 1);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfDocument.Load(bytes).Reader.TaggedContent());
        PdfStructureElementInfo ancestor = Assert.Single(tagged.StructureElements, element => element.StructureType == "Art");
        string source = PdfEncoding.Latin1GetString(bytes);
        System.Text.RegularExpressions.Match ancestorObject = System.Text.RegularExpressions.Regex.Match(
            source, "(?ms)^" + ancestor.ObjectNumber + " 0 obj\\s*(?<dictionary><<.*?>>)\\s*endobj$");

        Assert.True(ancestorObject.Success);
        Assert.DoesNotContain("/Pg", ancestorObject.Groups["dictionary"].Value, StringComparison.Ordinal);
    }

    [Fact]
    public void CrossPageFigureSemanticsEmitAlternateTextOnce() {
        byte[] bytes = PdfDocument.Create(document => document
            .Settings(options => {
                options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
                options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(180));
                options.PageSetup.Margins.Set(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(24)));
            })
            .Content(content => content.Semantic(PdfSemanticRole.Figure, figure => {
                for (int index = 0; index < 24; index++) {
                    figure.AddParagraph("Figure narrative " + index);
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
                .Content(nested => nested.Deferred(_ => deferred => deferred.AddParagraph("Page-aware content"))))))
            .ToBytes());

        Assert.Contains("height can be determined before rendering", exception.Message, StringComparison.Ordinal);
    }

    private sealed class TestComponent : IPdfComponent {
        public void Compose(PdfContentBuilder content) {
            content.AddParagraph("Reusable element content");
        }
    }
}
