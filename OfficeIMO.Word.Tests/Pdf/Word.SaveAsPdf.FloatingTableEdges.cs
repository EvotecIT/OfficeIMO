using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    // Positions measured from Word's own PDF export of the same layouts.
    [Theory]
    [InlineData(180, 0, 189.74)]
    [InlineData(0, 0, 181.22)]
    [InlineData(180, 4, 189.50)]
    [InlineData(180, 12, 190.61)]
    [InlineData(0, 24, 183.50)]
    public void SaveAsPdf_TextBesideFloatingTableStartsPastItsOutsideEdge(int rightFromText, int borderSize, double expectedLeft) {
        using WordDocument document = CreateFloatingEdgeDocument();
        AddFloatingEdgeTable(document, new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Text,
            TablePositionXAlignment = HorizontalAlignmentValues.Left, TablePositionY = 1, RightFromText = (short)rightFromText
        }, borderSize);
        document.AddParagraph(FloatingEdgeWords(60));

        using PdfPigDocument pdf = RenderFloatingEdgePdf(document);
        Assert.Equal(expectedLeft, FindFloatingEdgeWord(pdf, "w01").BoundingBox.Left, 0.15);
    }

    [Theory]
    [InlineData(0, 436.75)]
    [InlineData(12, 436.75)]
    public void SaveAsPdf_RightAlignedFloatingTablePlacesItsOutsideEdgeAtTheMargin(int borderSize, double expectedCellText) {
        using WordDocument document = CreateFloatingEdgeDocument();
        AddFloatingEdgeTable(document, new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Text,
            TablePositionXAlignment = HorizontalAlignmentValues.Right, TablePositionY = 1
        }, borderSize);
        document.AddParagraph(FloatingEdgeWords(60));

        using PdfPigDocument pdf = RenderFloatingEdgePdf(document);
        Assert.Equal(expectedCellText, FindFloatingEdgeWord(pdf, "FLOATCELL").BoundingBox.Left, 0.15);
    }

    [Fact]
    public void SaveAsPdf_BorderedFloatingTableStartsItsOutsideBorderAtTheAnchor() {
        using WordDocument document = CreateFloatingEdgeDocument();
        AddFloatingEdgeTable(document, new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Text,
            TablePositionXAlignment = HorizontalAlignmentValues.Left, TablePositionY = 1
        }, borderSize: 24);
        document.AddParagraph(FloatingEdgeWords(60));

        using PdfPigDocument pdf = RenderFloatingEdgePdf(document);
        Assert.Equal(78.86, FindFloatingEdgeWord(pdf, "FLOATCELL").BoundingBox.Left, 0.15);
    }

    [Fact]
    public void SaveAsPdf_NarrowRoomBesideAFloatingTableStaysEmpty() {
        // The 18pt gap before this table stays empty; text starts past its outside edge.
        using WordDocument document = CreateFloatingEdgeDocument();
        AddFloatingEdgeTable(document, new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Text,
            TablePositionX = 360, TablePositionY = 1
        }, borderSize: 0);
        document.AddParagraph(FloatingEdgeWords(80));

        using PdfPigDocument pdf = RenderFloatingEdgePdf(document);
        Assert.Equal(199.25, FindFloatingEdgeWord(pdf, "w01").BoundingBox.Left, 0.15);
    }

    private static WordDocument CreateFloatingEdgeDocument() {
        WordDocument document = WordDocument.Create();
        document.Sections[0].PageSetup.Size.PageSize = WordPageSize.Letter;
        document.Sections[0].PageSetup.Margins.Left = document.Sections[0].PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points(72D);
        document.AddParagraph("Intro.");
        return document;
    }

    private static void AddFloatingEdgeTable(WordDocument document, TablePositionProperties position, int borderSize) {
        WordTable table = document.AddTable(1, 1);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(108D));
        table.Rows[0].Cells[0].Paragraphs[0].Text = "FLOATCELL";
        TableProperties properties = table._tableProperties!;
        properties.TablePositionProperties = position;
        properties.TableStyle?.Remove();
        properties.TableBorders?.Remove();
        if (borderSize > 0) {
            properties.TableBorders = new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = (uint)borderSize, Space = 0, Color = "auto" },
                new LeftBorder { Val = BorderValues.Single, Size = (uint)borderSize, Space = 0, Color = "auto" },
                new BottomBorder { Val = BorderValues.Single, Size = (uint)borderSize, Space = 0, Color = "auto" },
                new RightBorder { Val = BorderValues.Single, Size = (uint)borderSize, Space = 0, Color = "auto" });
        }
    }

    private static string FloatingEdgeWords(int count) =>
        string.Join(" ", Enumerable.Range(1, count).Select(index => "w" + index.ToString("00", System.Globalization.CultureInfo.InvariantCulture)));

    private static PdfPigDocument RenderFloatingEdgePdf(WordDocument document) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        return PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
    }

    private static UglyToad.PdfPig.Content.Word FindFloatingEdgeWord(PdfPigDocument pdf, string token) =>
        pdf.GetPages().SelectMany(page => page.GetWords()).First(word => word.Text == token);
}
