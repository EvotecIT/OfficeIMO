using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunningContentTests {
    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 2)]
    public void ImportedFooterFirstLineOverlapRetainsHeaderWithoutChangingOrdinaryFlow(bool allowOverlap, int pages) {
        var options = Options();
        options.PageHeight = 300; options.MarginTop = options.MarginBottom = 72;
        var document = PdfDocument.Create(options);
        document.Header(header => header.Text("FALLBACK").Content(content => content.Spacer(96), 36));
        document.Footer(footer => footer.Content(content => content.Spacer(96), 36));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        document.Paragraph(paragraph => paragraph.Text("ALPHA").LineBreak().Text("BETA"), style: new PdfParagraphStyle {
            FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(48), SpacingBefore = 0, SpacingAfter = 0,
            WidowControl = false, AllowFirstLineOverlapWithRunningFooter = allowOverlap
        });

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.Contains("ALPHA", pdf.GetPage(1).Text);
        if (allowOverlap) {
            Assert.DoesNotContain("FALLBACK", pdf.GetPage(1).Text);
            Assert.DoesNotContain("BETA", pdf.GetPage(1).Text);
            Assert.Contains("BETA", pdf.GetPage(2).Text);
        } else {
            Assert.Contains("FALLBACK", pdf.GetPage(1).Text);
            Assert.Contains("BETA", pdf.GetPage(1).Text);
        }
    }

    [Fact]
    public void EmptyOrdinaryRunningContentDoesNotExpandTheBodyFrame() {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Content(_ => { }, 80, 20));
        document.Footer(footer => footer.Content(_ => { }, 60, 20));
        document.Paragraph(paragraph => paragraph.FontSize(40)
            .Text("A").LineBreak().Text("A").LineBreak().Text("A").LineBreak().Text("A"));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Equal(4, pdf.GetPage(1).Letters.Count(letter => letter.Value == "A"));
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public void ImportedHeaderFitReservesActiveContainerPaddingAcrossExplicitBreaks(bool breakEndsPanel, bool repeatDecoration) {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Text("TEXT")
            .Content(content => content.Table(new[] { new[] { "TALL" } }, style: TableStyle(230)), 18)
            .FirstPageContent(content => content.Table(new[] { new[] { "FIRST HEADER" } }, style: TableStyle(24)), 18));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        document.Panel(panel => {
            panel.Paragraph(paragraph => paragraph.Text("FIRST")).PageBreak();
            if (!breakEndsPanel) panel.Paragraph(paragraph => paragraph.FontSize(20).Text("BODY"));
        }, new PdfPanelStyle { PaddingX = 12, PaddingY = 20, RepeatFragmentDecoration = repeatDecoration, SpacingBefore = 0, SpacingAfter = 0 });
        if (breakEndsPanel) document.Paragraph(paragraph => paragraph.FontSize(20).Text("BODY"));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FIRST", pdf.GetPage(1).Text);
        Assert.Contains("TEXT", pdf.GetPage(2).Text);
        Assert.Contains("BODY", pdf.GetPage(2).Text);
        Assert.DoesNotContain("TALL", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedHeaderFitReservesActiveContainerBottomInsetForContinuationLines(bool nested) {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Text("TEXT").Content(content =>
            content.Table(new[] { new[] { "TALL" } }, style: TableStyle(230)), 18));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        var style = new PdfPanelStyle { PaddingX = 12, PaddingY = 0, FragmentBottomInset = nested ? 10 : 20, SpacingBefore = 0, SpacingAfter = 0 };
        document.Panel(panel => {
            if (nested) panel.Panel(inner => inner.Paragraph(paragraph =>
                paragraph.Text("SMALL").LineBreak().FontSize(20).Text("LARGE")), style);
            else panel.Paragraph(paragraph => paragraph.Text("SMALL").LineBreak().FontSize(20).Text("LARGE"));
        }, style);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("TALL", pdf.GetPage(1).Text);
        Assert.Contains("SMALL", pdf.GetPage(1).Text);
        Assert.Contains("TEXT", pdf.GetPage(2).Text);
        Assert.Contains("LARGE", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ImportedHeaderFitSkipsLeadingNonvisualContent(int leadingKind) {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Text("TEXT").Content(content =>
            content.Table(new[] { new[] { "FLOW" } }, style: TableStyle(246)), 18));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        if (leadingKind == 2) {
            document.Page(page => page.Content(content => content.Panel(panel => panel
                .Semantic(PdfSemanticRole.Division, empty => { })
                .Paragraph(paragraph => paragraph.FontSize(40).Text("BODY")),
                new PdfPanelStyle { PaddingX = 0, PaddingY = 2, SpacingBefore = 0, SpacingAfter = 0 })));
        } else if (leadingKind == 1) {
            document.Page(page => page.Content(content => content
                .Semantic(PdfSemanticRole.Division, empty => { })
                .Paragraph(paragraph => paragraph.FontSize(40).Text("BODY"))));
        } else {
            document.Bookmark("body").Paragraph(paragraph => paragraph.FontSize(40).Text("BODY"));
        }
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("TEXT", pdf.GetPage(1).Text);
        Assert.Contains("BODY", pdf.GetPage(1).Text);
    }

    [Fact]
    public void ImportedHeaderFitUsesTheActualParagraphContinuationLine() {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Text("TEXT").Content(content =>
            content.Table(new[] { new[] { "FLOW" } }, style: TableStyle(246)), 18));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        document.Paragraph(paragraph => paragraph.Text("SMALL").LineBreak().FontSize(40).Text("LARGE"));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FLOW", pdf.GetPage(1).Text);
        Assert.Contains("SMALL", pdf.GetPage(1).Text);
        Assert.Contains("TEXT", pdf.GetPage(2).Text);
        Assert.Contains("LARGE", pdf.GetPage(2).Text);
    }

    [Fact]
    public void ImportedTextFallbackIsLocalToThePageVariantAndReservesTheBodyLine() {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Text("DEFAULT TEXT").FirstPageText("FIRST TEXT")
            .Content(content => content.Table(new[] { new[] { "FLOW" } }, style: TableStyle(24)), 18)
            .FirstPageContent(content => content.Table(new[] { new[] { "TALL" } }, style: TableStyle(246)), 18));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        document.Paragraph(paragraph => paragraph.FontSize(40).Text("BODY")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("BODY"));

        for (int pass = 0; pass < 2; pass++) {
            using var pdf = PdfPigDocument.Open(document.ToBytes());
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("FIRST TEXT", pdf.GetPage(1).Text);
            Assert.Contains("BODY", pdf.GetPage(1).Text);
            Assert.DoesNotContain("TALL", pdf.GetPage(1).Text);
            Assert.Contains("FLOW", pdf.GetPage(2).Text);
            Assert.Contains("BODY", pdf.GetPage(2).Text);
            Assert.DoesNotContain("DEFAULT TEXT", pdf.GetPage(2).Text);
        }
    }

    private static PdfOptions Options(long memoryLimit = 1024 * 1024) => new() {
        PageWidth = 240, PageHeight = 320, MarginTop = 32, MarginBottom = 32,
        MarginLeft = 24, MarginRight = 24, DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 11, PageContentMemoryLimitBytes = memoryLimit
    };

    private static PdfTableStyle TableStyle(double height) {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 100 };
        style.FixedRowHeights = new List<double?> { height };
        style.CellPaddingX = 3; style.CellPaddingY = 3;
        style.BorderWidth = 1;
        return style;
    }

    [Theory]
    [InlineData(0L)]
    [InlineData(1048576L)]
    public void Story_tables_preserve_rotated_text_links_and_shared_resource_lifetime(long memoryLimit) {
        var document = PdfDocument.Create(Options(memoryLimit));
        var cell = new PdfTableCell(new[] { PdfTextRun.Normal("HEAD") }, paragraphs: null,
            linkUri: "https://example.com/running-header", textRotation: -90);
        document.Header(header => header.Content(content => content.Table(new[] { new[] { cell } }, style: TableStyle(80)), 18, 6));
        document.Paragraph(paragraph => paragraph.Text("BODY")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("BODY"));
        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (var page in pdf.GetPages()) {
            Assert.Contains("HEAD", page.Text);
            Assert.Contains("BODY", page.Text);
            var head = page.Letters.First(letter => letter.Value == "H");
            Assert.True(Math.Abs(head.EndBaseLine.Y - head.StartBaseLine.Y) > 1);
            var body = page.Letters.First(letter => letter.Value == "B");
            Assert.True(body.BoundingBox.Top < 320 - 18 - 80 - 6);
        }
        Assert.Equal(2, PdfInspector.Inspect(bytes).GetAnnotationsBySubtype("Link").Count());
    }

    [Fact]
    public void Selected_variants_reserve_their_own_height_and_materialize_final_page_fields() {
        var document = PdfDocument.Create(Options());
        Action<PdfContentBuilder> Compose(PdfRunningContentContext context, string label, double height) =>
            content => content.Table(new[] { new[] { $"{label}{context.PageNumber}/{context.DocumentPages}" } }, style: TableStyle(height));
        document.Header(header => header.Content(context => Compose(context, "Default", 120), 18, 6)
            .FirstPageContent(context => Compose(context, "First", 24), 18, 6)
            .EvenPagesContent(context => Compose(context, "Even", 72), 18, 6));
        document.Footer(footer => footer.Content(context => content => content.Paragraph(paragraph => paragraph.Text($"F{context.PageNumber}/{context.DocumentPages}")), 18)
            .FirstPageContent(context => content => content.Paragraph(paragraph => paragraph.Text($"F{context.PageNumber}/{context.DocumentPages}")), 18)
            .EvenPagesContent(context => content => content.Paragraph(paragraph => paragraph.Text($"F{context.PageNumber}/{context.DocumentPages}")), 18));
        document.Paragraph(paragraph => paragraph.Text("BODY")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("BODY")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("BODY"));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        var labels = new[] { "First1/3", "Even2/3", "Default3/3" };
        var bodyTops = new double[3];
        for (int index = 0; index < 3; index++) {
            var page = pdf.GetPage(index + 1);
            Assert.Contains(labels[index], page.Text);
            Assert.Contains($"F{index + 1}/3", page.Text);
            bodyTops[index] = page.Letters.First(letter => letter.Value == "B").BoundingBox.Top;
        }
        Assert.InRange(bodyTops[0] - bodyTops[1], 47.99, 48.01);
        Assert.InRange(bodyTops[0] - bodyTops[2], 95.99, 96.01);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Physical_section_counts_are_independent_of_visible_numbering_sequences(bool continuingSections) {
        var document = PdfDocument.Create(Options());
        void AddSection(string label, int? start) => document.Section(section => {
            if (start.HasValue) section.PageNumberStart(start.Value);
            section.Header(header => header.Content(context => content => content.Paragraph(paragraph =>
                paragraph.Text($"{label}/{context.PageNumber}/{context.TotalPages}/{context.SectionPageNumber}/{context.SectionPages}/{context.DocumentPages}")), 18));
            section.Footer(footer => footer.Text("Legacy {page}/{pages}"));
            section.Content(content => content.Paragraph(paragraph => paragraph.Text("BODY"))
                .PageBreak().Paragraph(paragraph => paragraph.Text("BODY")));
        });
        AddSection("First", continuingSections ? null : 5);
        if (continuingSections) AddSection("Second", null);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(continuingSections ? 4 : 2, pdf.NumberOfPages);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            int visible = continuingSections ? page : page + 4;
            int total = continuingSections ? 4 : 6;
            string label = page <= 2 ? "First" : "Second";
            Assert.Contains($"{label}/{visible}/{total}/{(page - 1) % 2 + 1}/2/{pdf.NumberOfPages}", pdf.GetPage(page).Text);
            Assert.Contains($"Legacy {visible}/{total}", pdf.GetPage(page).Text);
        }
    }

    [Fact]
    public void Header_only_document_renders_one_page_and_later_text_replaces_rich_content() {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Content(content => content.Paragraph(paragraph => paragraph.Text("RICH")), 18));
        using (var pdf = PdfPigDocument.Open(document.ToBytes())) {
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("RICH", pdf.GetPage(1).Text);
        }
        document.Header(header => header.Text("TEXT"));
        document.Paragraph(paragraph => paragraph.Text("BODY"));
        using var replacement = PdfPigDocument.Open(document.ToBytes());
        Assert.Contains("TEXT", replacement.GetPage(1).Text);
        Assert.DoesNotContain("RICH", replacement.GetPage(1).Text);
    }

    [Fact]
    public void Running_page_breaks_are_rejected_instead_of_changing_document_pagination() {
        var document = PdfDocument.Create(Options());
        document.Header(header => header.Content(content => content.Paragraph(paragraph => paragraph.Text("HEAD")).PageBreak(), 18));
        document.Paragraph(paragraph => paragraph.Text("BODY"));
        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Story_only_embedded_glyphs_reach_font_subsets_and_unicode_maps(bool namedFamily) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        var family = new PdfEmbeddedFontFamily("Story Font", font);
        PdfOptions options = Options(0);
        if (namedFamily) options.RegisterNamedFontFamily(family);
        else options.UseFontFamily(family);
        var document = PdfDocument.Create(options);
        document.Header(header => header.Content(context => content => content.Paragraph(paragraph => {
            if (namedFamily) paragraph.FontFamily("Story Font");
            paragraph.Text($"HEADER{context.PageNumber}/{context.DocumentPages}");
        }), 18));
        document.Paragraph(paragraph => paragraph.Text("body")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("body"));
        // Each serialization owns/reset its generation usage; a replay must retain story-only glyphs.
        for (int replay = 0; replay < 2; replay++) {
            using var pdf = PdfPigDocument.Open(document.ToBytes());
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("HEADER1/2", pdf.GetPage(1).Text);
            Assert.Contains("HEADER2/2", pdf.GetPage(2).Text);
        }
    }

    [Theory]
    [InlineData(0L)]
    [InlineData(1048576L)]
    public void Story_images_and_transparency_groups_survive_borrowed_store_disposal(long memoryLimit) {
        var shape = OfficeShape.Rectangle(20, 20);
        shape.FillColor = OfficeColor.Red;
        var source = new OfficeDrawing(20, 20).AddShape(shape, 0, 0);
        var drawing = new OfficeDrawing(20, 20).AddEffectDrawing(source, OfficeTransform.Identity, .5);
        byte[] image = PdfPngTestImages.CreateRgbPng(1, 1);
        var document = PdfDocument.Create(Options(memoryLimit));
        document.Header(header => header.Content(content => content.Drawing(drawing, spacingBefore: 0, spacingAfter: 0), 18));
        document.Footer(footer => footer.Content(content => content.Image(image, 12, 12), 18));
        document.Paragraph(paragraph => paragraph.Text("BODY")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("BODY"));
        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (var page in pdf.GetPages()) {
            Assert.Contains("BODY", page.Text);
            Assert.Single(page.GetImages());
        }
        string serialized = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Group << /S /Transparency", serialized);
        Assert.Contains("/ca 0.5 /CA 0.5", serialized);
    }
}
