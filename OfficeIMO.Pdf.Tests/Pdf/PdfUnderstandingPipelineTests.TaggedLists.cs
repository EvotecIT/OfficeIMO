using OfficeIMO.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfUnderstandingPipelineTests {
    [Fact]
    public void StructuredRead_AssociatesNestedParagraphAndLinkWithTheirTaggedListItem() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf()).Read();
        PdfLogicalPage page = Assert.Single(result.Pages);

        Assert.Equal(2, page.ListItems.Count);
        PdfLogicalListItem first = page.ListItems[0];
        Assert.Equal("1.", first.Marker);
        Assert.Equal("Linked incident title Incident detail", first.Text);
        Assert.DoesNotContain("Neighboring incident", first.Text, StringComparison.Ordinal);
        Assert.Equal("2.", page.ListItems[1].Marker);
        Assert.Equal("Neighboring incident", page.ListItems[1].Text);
        Assert.Contains(page.Paragraphs, paragraph => paragraph.Text.IndexOf("Linked incident title", StringComparison.Ordinal) >= 0);
        Assert.All(first.Lines, line => Assert.Contains(line, page.TextBlocks));
        Assert.DoesNotContain(first.Runs, static run => run.SourceSpan?.MarkedContentId == 0);
        Assert.Contains(first.Runs, static run => run.SourceSpan?.MarkedContentId == 1);
        Assert.Equal(1, CountText(result.ToMarkdown(), "Linked incident title"));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Incident detail"));
    }

    [Fact]
    public void StructuredRead_KeepsNestedListBodiesSeparateAndRetainsTheirLevels() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf(nested: true)).Read();
        PdfLogicalPage page = Assert.Single(result.Pages);
        Assert.Equal(new[] { 1, 2 }, page.ListItems.Select(static item => item.Level));
        Assert.Equal("Linked incident title Incident detail", page.ListItems[0].Text);
        Assert.Equal("Neighboring incident", page.ListItems[1].Text);
        Assert.Equal(1, CountText(result.ToMarkdown(), "Neighboring incident"));
    }

    [Fact]
    public void StructuredRead_PreservesHeadingProjectionInsideAListAssociation() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf(heading: true)).Read();
        PdfLogicalPage page = Assert.Single(result.Pages);
        Assert.Equal("Linked incident title Incident detail", page.ListItems[0].Text);
        PdfLogicalHeading heading = Assert.Single(page.Headings);
        Assert.Equal("Linked incident title", heading.Text);
        Assert.Equal(2, heading.Level);
        Assert.Contains(PdfLogicalReadingOrderAnalysis.Analyze(page), static item => item.Kind == PdfLogicalReadingOrderKind.Heading);
        Assert.Equal(1, CountText(result.ToMarkdown(), "Linked incident title"));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Incident detail"));
    }

    [Fact]
    public void StructuredRead_DoesNotDuplicateListTextWhenAParagraphAlsoContainsUntaggedText() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf(trailing: true)).Read();
        Assert.Equal(2, Assert.Single(result.Pages).ListItems.Count);
        string markdown = result.ToMarkdown();
        Assert.Equal(1, CountText(markdown, "Neighboring incident"));
        Assert.Equal(1, CountText(markdown, "Unrelated paragraph"));
    }

    [Fact]
    public void StructuredRead_ReconstructsListAssociationsForEachRepeatedPageSelection() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf()).Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(1, 1)
        });
        Assert.Equal(2, result.Pages.Count);
        Assert.All(result.Pages, page => {
            Assert.Equal(2, page.ListItems.Count);
            Assert.Equal("Linked incident title Incident detail", page.ListItems[0].Text);
        });
        Assert.NotSame(result.Pages[0].ListItems[0], result.Pages[1].ListItems[0]);
        Assert.NotSame(result.Pages[0].ListItems[0].Line, result.Pages[1].ListItems[0].Line);
    }

    [Fact]
    public void StructuredRead_RetainsEveryPageFragmentOfOneTaggedListItem() {
        const string phrase = "Continuation text keeps one logical list item";
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            CompressContentStreams = false, PageWidth = 170, PageHeight = 150,
            MarginLeft = 24, MarginRight = 24, MarginTop = 24, MarginBottom = 24, DefaultFontSize = 10
        }).TaggedPdfCatalogMarkers().Bullets(new[] { string.Join(" ", Enumerable.Repeat(phrase, 18)) }).ToBytes();
        PdfDocumentReadResult result = PdfDocument.Load(bytes).Read();
        Assert.True(result.Pages.Count > 1);
        Assert.All(result.Pages, page => Assert.Single(page.ListItems));
        Assert.Equal(string.Join(" ", Enumerable.Repeat(phrase, 18)), string.Join(" ", result.ListItems.Select(static item => item.Text)));
        Assert.Equal(18, CountText(result.ToMarkdown(), "Continuation"));
    }

    private static int CountText(string text, string value) => text.Split(new[] { value }, StringSplitOptions.None).Length - 1;

    [Fact]
    public void StructuredRead_PreservesTableProjectionInsideAListAssociation() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }).TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.List, list => list
                .Structure(PdfCanvasStructureRole.ListItem, item => item
                    .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("1.", 20D, 80D, 20D, 16D))
                    .Structure(PdfCanvasStructureRole.ListBody, body => body
                        .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Incident table", 50D, 80D, 180D, 16D))
                        .Structure(PdfCanvasStructureRole.Table, table => table
                            .Structure(PdfCanvasStructureRole.TableRow, row => row
                                .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Metric", 50D, 110D, 80D, 16D))
                                .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Value", 160D, 110D, 80D, 16D)))
                            .Structure(PdfCanvasStructureRole.TableRow, row => row
                                .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("Quality", 50D, 130D, 80D, 16D))
                                .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("High", 160D, 130D, 80D, 16D))))))))
            .ToBytes();
        PdfDocumentReadResult result = PdfDocument.Load(bytes).Read();
        PdfLogicalPage page = Assert.Single(result.Pages);
        PdfLogicalListItem item = Assert.Single(page.ListItems);
        Assert.Contains("Incident table", item.Text, StringComparison.Ordinal);
        Assert.Contains("Quality", item.Text, StringComparison.Ordinal);
        Assert.Single(page.Tables);
        Assert.Contains(PdfLogicalReadingOrderAnalysis.Analyze(page), static candidate => candidate.Kind == PdfLogicalReadingOrderKind.Table);
        Assert.Equal(1, CountText(result.ToMarkdown(), "Quality"));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Incident table"));
        AssertConsumerTextOccursOnce(result, "Incident table", "Quality");
        Assert.DoesNotContain(PdfReaderAdapter.ReadDocument(result).Blocks,
            static block => block.Kind == "list-item" && block.Text.IndexOf("Quality", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void StructuredRead_AssociatesSeparateItemsSharingOnePhysicalLineWithoutDuplicatingExport() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }).TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.List, list => list
                .Structure(PdfCanvasStructureRole.ListItem, item => item
                    .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("1.", 20D, 100D, 20D, 16D))
                    .Structure(PdfCanvasStructureRole.ListBody, body => body
                        .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Left incident", 50D, 100D, 150D, 16D))))
                .Structure(PdfCanvasStructureRole.ListItem, item => item
                    .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("2.", 250D, 100D, 20D, 16D))
                    .Structure(PdfCanvasStructureRole.ListBody, body => body
                        .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Right incident", 280D, 100D, 150D, 16D))))))
            .ToBytes();
        PdfDocumentReadResult result = PdfDocument.Load(bytes).Read();
        Assert.Equal(new[] { "Left incident", "Right incident" }, result.ListItems.Select(static item => item.Text));
        Assert.Equal(new[] { "1.", "2." }, result.ListItems.Select(static item => item.Marker));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Left incident"));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Right incident"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TaggedListConsumers_ProjectParagraphBodiesOnceWithEitherOrderingMode(bool sharedOrder) {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf()).Read();
        AssertConsumerTextOccursOnce(result, sharedOrder, "Linked incident title", "Incident detail", "Neighboring incident");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TaggedListConsumers_RetainHeadingRoleWithoutDuplicatingItsListAssociation(bool sharedOrder) {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf(heading: true)).Read();
        AssertConsumerTextOccursOnce(result, sharedOrder, "Linked incident title", "Incident detail", "Neighboring incident");
        Assert.Contains(PdfReaderAdapter.ReadDocument(result).Blocks,
            static block => block.Kind == "heading" && block.Text == "Linked incident title");
    }

    [Fact]
    public void WordImportWithoutLists_RetainsTaggedParagraphBodies() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateParagraphTaggedListPdf()).Read();
        using OfficeIMO.Word.WordDocument word = result.ToWordDocument(new PdfToWordOptions { ImportLists = false });
        string text = ReadWordBody(word);
        Assert.Equal(1, CountText(text, "Linked incident title"));
        Assert.Equal(1, CountText(text, "Incident detail"));
        Assert.Equal(1, CountText(text, "Neighboring incident"));
    }

    [Fact]
    public void TaggedListRuns_RetainTheExactOwnerOfIdenticalBodiesSharingOneLine() {
        PdfDocumentReadResult result = PdfDocument.Load(CreateSharedLineTaggedListPdf()).Read();
        PdfLogicalListItem[] items = result.ListItems.ToArray();
        Assert.Equal(2, items.Length);
        Assert.Same(items[0].Line, items[1].Line);
        Assert.Equal(new[] { "Same", "Same" }, items.Select(static item => item.Text));
        for (int index = 0; index < items.Length; index++) {
            PdfLogicalTextRun run = Assert.Single(items[index].Runs);
            Assert.Same(items[index].Line.Spans.Single(span => span.MarkedContentId == 1 + 2 * index), run.SourceSpan);
        }
        AssertConsumerTextOccursTwice(result, "Same");
    }

    [Fact]
    public void TaggedListRuns_RetainDiscontiguousBodySourcesAroundUnownedText() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }).TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.List, list => list
                    .Structure(PdfCanvasStructureRole.ListItem, item => item
                        .Structure(PdfCanvasStructureRole.ListBody, body => body
                            .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph
                                .Text("Left", 50D, 100D, 35D, 16D)
                                .Text("Tail", 130D, 100D, 35D, 16D)))))
                .Text("Other", 90D, 100D, 35D, 16D))
            .ToBytes();
        PdfDocumentReadResult result = PdfDocument.Load(bytes).Read();
        PdfLogicalListItem item = Assert.Single(result.ListItems);
        Assert.Equal("Left Tail", item.Text);
        Assert.Equal(item.Text, string.Concat(item.Runs.Select(static run => run.Text)));
        Assert.Same(item.Line.Spans.Single(static span => span.Text == "Left"),
            Assert.Single(item.Runs, static run => run.Text == "Left").SourceSpan);
        Assert.Same(item.Line.Spans.Single(static span => span.Text == "Tail"),
            Assert.Single(item.Runs, static run => run.Text == "Tail").SourceSpan);
        Assert.DoesNotContain(item.Runs, static run => run.SourceSpan?.Text == "Other");
        AssertConsumerTextOccursOnce(result, "Left", "Other", "Tail");
    }

    private static byte[] CreateSharedLineTaggedListPdf() => PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
        .TaggedPdfCatalogMarkers().Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.List, list => list
            .Structure(PdfCanvasStructureRole.ListItem, item => item
                .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("1.", 20D, 100D, 20D, 16D))
                .Structure(PdfCanvasStructureRole.ListBody, body => body
                    .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Same", 50D, 100D, 40D, 16D))))
            .Structure(PdfCanvasStructureRole.ListItem, item => item
                .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("2.", 100D, 100D, 20D, 16D))
                .Structure(PdfCanvasStructureRole.ListBody, body => body
                    .Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph.Text("Same", 130D, 100D, 40D, 16D))))))
        .ToBytes();

    private static void AssertConsumerTextOccursOnce(PdfDocumentReadResult result, params string[] values) =>
        AssertConsumerTextOccursOnce(result, true, values);

    private static void AssertConsumerTextOccursOnce(PdfDocumentReadResult result, bool sharedOrder, params string[] values) {
        string reader = string.Join(" ", PdfReaderAdapter.ReadDocument(result).Blocks.Select(static block => block.Text));
        string html = result.ToHtml(new PdfToHtmlOptions { UseSharedPageReadingOrder = sharedOrder });
        using OfficeIMO.Word.WordDocument word = result.ToWordDocument(new PdfToWordOptions { UseSharedPageReadingOrder = sharedOrder });
        string wordText = ReadWordBody(word);
        foreach (string value in values) {
            // Table content is represented in Reader tables, not duplicated in document blocks.
            if (result.Pages.All(page => page.Tables.Count == 0) || value == "Incident table") Assert.Equal(1, CountText(reader, value));
            Assert.Equal(1, CountText(html, value));
            Assert.Equal(1, CountText(wordText, value));
        }
    }

    private static void AssertConsumerTextOccursTwice(PdfDocumentReadResult result, string value) {
        Assert.Equal(2, CountText(string.Join(" ", PdfReaderAdapter.ReadDocument(result).Blocks.Select(static block => block.Text)), value));
        Assert.Equal(2, CountText(result.ToHtml(), value));
        using OfficeIMO.Word.WordDocument word = result.ToWordDocument();
        Assert.Equal(2, CountText(ReadWordBody(word), value));
    }

    private static string ReadWordBody(OfficeIMO.Word.WordDocument word) {
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
        return string.Join(" ", package.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Select(static paragraph => string.Concat(paragraph.Descendants<Text>().Select(static text => text.Text))));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StructuredRead_DistinguishesListOwnersWhenPageAndFormReuseAnMcid(bool typedReference) {
        string pageContent = "/LI << /MCID 0 >> BDC BT /F1 12 Tf 72 700 Td (Page incident) Tj ET EMC /Fx1 Do\n";
        string formContent = "/LI << /MCID 0 >> BDC BT /F1 12 Tf 72 500 Td (Form incident) Tj ET EMC\n";
        string type = typedReference ? "/Type /MCR " : string.Empty;
        byte[] bytes = BuildClassicPdf(
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 7 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /StructParents 0 /Resources << /Font << /F1 5 0 R >> /XObject << /Fx1 6 0 R >> >> /Contents 4 0 R >>",
            BuildStreamBody(string.Empty, pageContent),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            BuildStreamBody("/Type /XObject /Subtype /Form /BBox [0 0 612 792] /StructParent 1 /Resources << /Font << /F1 5 0 R >> >>", formContent),
            "<< /Type /StructTreeRoot /K [8 0 R 9 0 R] /ParentTree 10 0 R /ParentTreeNextKey 2 /RoleMap << /Entry /LI >> >>",
            "<< /Type /StructElem /S /Entry /P 7 0 R /Pg 3 0 R /K << " + type + "/Pg 3 0 R /Stm 4 0 R /MCID 0 >> >>",
            "<< /Type /StructElem /S /LI /P 7 0 R /Pg 3 0 R /K << " + type + "/Pg 3 0 R /Stm 6 0 R /MCID 0 >> >>",
            "<< /Nums [0 [8 0 R] 1 [9 0 R]] >>");
        PdfDocumentReadResult result = PdfDocument.Load(bytes).Read();
        Assert.Equal(new[] { "Page incident", "Form incident" }, result.ListItems.Select(static item => item.Text));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Page incident"));
        Assert.Equal(1, CountText(result.ToMarkdown(), "Form incident"));
    }

    private static byte[] CreateParagraphTaggedListPdf(bool nested = false, bool heading = false, bool trailing = false) {
        string content = "/Lbl << /MCID 0 >> BDC BT /F1 12 Tf 50 " + (heading ? "710" : "700") + " Td (1.) Tj ET EMC\n" +
            "/Link << /MCID 1 >> BDC BT /F1 12 Tf 80 700 Td (Linked incident title) Tj ET EMC\n" +
            "/P << /MCID 2 >> BDC BT /F1 12 Tf 80 680 Td (Incident detail) Tj ET EMC\n" +
            "/Lbl << /MCID 3 >> BDC BT /F1 12 Tf 50 640 Td (2.) Tj ET EMC\n" +
            "/P << /MCID 4 >> BDC BT /F1 12 Tf 80 640 Td (Neighboring incident) Tj ET EMC\n" +
            (trailing ? "BT /F1 12 Tf 80 620 Td (Unrelated paragraph) Tj ET\n" : string.Empty);
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 6 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /StructParents 0 >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "endstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 18 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /L /P 6 0 R /Pg 3 0 R /K [8 0 R " + (nested ? "" : "14 0 R") + "] >>",
            "<< /Type /StructElem /S /LI /P 7 0 R /K [9 0 R 10 0 R] >>",
            "<< /Type /StructElem /S /Lbl /P 8 0 R /K 0 >>",
            "<< /Type /StructElem /S /LBody /P 8 0 R /K [11 0 R 13 0 R " + (nested ? "19 0 R" : "") + "] >>",
            "<< /Type /StructElem /S /" + (heading ? "H2" : "P") + " /P 10 0 R /K [12 0 R] >>",
            "<< /Type /StructElem /S /Link /P 11 0 R /K 1 >>",
            "<< /Type /StructElem /S /P /P 10 0 R /K 2 >>",
            "<< /Type /StructElem /S /LI /P " + (nested ? "19" : "7") + " 0 R /K [15 0 R 16 0 R] >>",
            "<< /Type /StructElem /S /Lbl /P 14 0 R /K 3 >>",
            "<< /Type /StructElem /S /LBody /P 14 0 R /K [17 0 R] >>",
            "<< /Type /StructElem /S /P /P 16 0 R /K 4 >>",
            "<< /Nums [0 [9 0 R 12 0 R 13 0 R 15 0 R 17 0 R]] >>"
        };
        if (nested) objects.Add("<< /Type /StructElem /S /L /P 10 0 R /K [14 0 R] >>");
        return BuildClassicPdf(objects.ToArray());
    }
}
