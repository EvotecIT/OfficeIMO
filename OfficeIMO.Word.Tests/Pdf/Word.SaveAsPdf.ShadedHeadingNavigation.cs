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
}
