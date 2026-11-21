using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, 1, false)]
    [InlineData(WordFileFormat.Doc, 1, false)]
    [InlineData(WordFileFormat.Docx, 3, false)]
    [InlineData(WordFileFormat.Docx, 1, true)]
    [InlineData(WordFileFormat.Doc, 1, true)]
    public void SaveAsPdf_ShadedHeadingKeepsOutlineDestinationTagAndLink(WordFileFormat format, int level, bool columns) {
        using WordDocument source = WordDocument.Create();
        if (columns) source.Sections[0].ColumnCount = 2;
        WordParagraph heading = source.AddParagraph("NavigationMarker ");
        heading.Style = level == 1 ? WordParagraphStyles.Heading1 : WordParagraphStyles.Heading3;
        heading.ShadingFillColor = OfficeColor.Parse("000080");
        heading.AddHyperLink("LinkMarker", new Uri("https://example.com/heading"), tooltip: "Heading target");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PdfOptions = new PdfCreateOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers }
        }).ToBytes();

        PdfOutlineItem outline = Assert.Single(PdfInspector.Inspect(bytes).Outlines);
        Assert.Equal("NavigationMarker LinkMarker", outline.Title);
        Assert.Equal(1, outline.PageNumber);
        Assert.Contains($"/Type /StructElem /S /H{level}", Encoding.UTF8.GetString(bytes), StringComparison.Ordinal);
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(bytes);
        Assert.NotEmpty(logical.NamedDestinations);
        Assert.Equal("Heading target", Assert.Single(logical.GetLinksByUri("https://example.com/heading")).Contents);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Doc, true)]
    public void SaveAsPdf_HeadingAppliesPositiveFirstLineIndent(WordFileFormat format, bool shaded) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].PageSetup.Margins.Left = OfficeLength.Points(40);
        WordParagraph heading = source.AddParagraph("IndentMarker");
        heading.Style = WordParagraphStyles.Heading1;
        if (shaded) heading.ShadingFillColor = OfficeColor.Parse("000080");
        heading.IndentationBefore = OfficeLength.Points(18);
        heading.IndentationFirstLine = OfficeLength.Points(9);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        // Margin 40 + left indent 18 + first-line indent 9.
        Assert.Equal(67, FindWordStartX(pdf.GetPage(1), "IndentMarker"), 2);
        Assert.Equal("IndentMarker", Assert.Single(PdfInspector.Inspect(bytes).Outlines).Title);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_ShadedHeadingPaintsItsFillBehindTheHeadingText(WordFileFormat format) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].PageSetup.Margins.Left = OfficeLength.Points(40);
        source.AddParagraph("BodyBeforeMarker");
        WordParagraph heading = source.AddParagraph("FillMarker");
        heading.Style = WordParagraphStyles.Heading1;
        heading.ShadingFillColor = OfficeColor.Parse("FFFF00");
        heading.IndentationBefore = OfficeLength.Points(18);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        var word = page.GetWords().Single(candidate => candidate.Text == "FillMarker");
        var fill = Assert.Single(page.Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D))
            .GetBoundingRectangle()!.Value;
        // The fill starts at the heading indent (margin 40 + indent 18) and covers the heading line.
        Assert.Equal(58D, fill.Left, 1);
        Assert.InRange(word.BoundingBox.Bottom, fill.Bottom, fill.Top);
        Assert.InRange(word.BoundingBox.Top, fill.Bottom, fill.Top);
        Assert.Equal("FillMarker", Assert.Single(PdfInspector.Inspect(bytes).Outlines).Title);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsPdf_ShadedHeadingStaysWithTheFollowingParagraph(bool keepWithNext) {
        using WordDocument source = WordDocument.Create();
        WordParagraph filler = source.AddParagraph("FillerMarker");
        filler.LineSpacing = WordLineSpacing.Exact(OfficeLength.Points(95));
        filler.LineSpacingBefore = filler.LineSpacingAfter = OfficeLength.Zero;
        WordParagraph heading = source.AddParagraph("KeptHeadingMarker");
        heading.Style = WordParagraphStyles.Heading1;
        heading.ShadingFillColor = OfficeColor.Parse("000080");
        heading.KeepWithNextOverride = keepWithNext;
        WordParagraph following = source.AddParagraph("FollowingMarker");
        following.LineSpacing = WordLineSpacing.Exact(OfficeLength.Points(40));
        following.KeepLinesTogetherOverride = true;
        byte[] bytes = source.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeLength.Points(400), OfficeLength.Points(220)),
            Margins = PdfPageMargins.Uniform(OfficeLength.Points(40))
        }).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        // The heading fits below the filler; only its keep with the following paragraph moves it.
        Assert.Equal(!keepWithNext, pdf.GetPage(1).Text.Contains("KeptHeadingMarker", StringComparison.Ordinal));
        Assert.Contains("FollowingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        if (keepWithNext) {
            // The heading's outline destination follows it to the page it moved to.
            Assert.Equal(2, Assert.Single(PdfInspector.Inspect(bytes).Outlines).PageNumber);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ShadedHeadingKeepsItsSpacingBeforeAtThePageTop(bool pageBreakBefore) {
        double HeadingBaseline(bool shaded) {
            using WordDocument source = WordDocument.Create();
            if (pageBreakBefore) source.AddParagraph("FirstPageMarker").AddBreak(WordBreakType.Page);
            WordParagraph heading = source.AddParagraph("TopHeadingMarker");
            heading.Style = WordParagraphStyles.Heading1;
            heading.LineSpacingBefore = OfficeLength.Points(30);
            if (shaded) heading.ShadingFillColor = OfficeColor.Parse("FFFF00");
            source.AddParagraph("BodyMarker");
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(source.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes());
            return pdf.GetPages().SelectMany(page => page.GetWords()).Single(word => word.Text == "TopHeadingMarker").Letters[0].StartBaseLine.Y;
        }

        Assert.Equal(HeadingBaseline(shaded: false), HeadingBaseline(shaded: true), 2);
    }
}
