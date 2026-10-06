using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, WordFileFormat.Docx)]
    [InlineData(false, false, WordFileFormat.Doc)]
    [InlineData(false, true, WordFileFormat.Docx)]
    [InlineData(false, true, WordFileFormat.Doc)]
    [InlineData(true, false, WordFileFormat.Docx)]
    [InlineData(true, false, WordFileFormat.Doc)]
    [InlineData(true, true, WordFileFormat.Docx)]
    [InlineData(true, true, WordFileFormat.Doc)]
    public void SaveAsPdf_PageGutterChangesTheBodyFrameInItsAuthoredDirection(bool atTop, bool onRight, WordFileFormat format) {
        var baseline = RenderGutterMarkers(atTop, onRight, format, 0);
        var gutter = RenderGutterMarkers(atTop, onRight, format, 400);
        Assert.Equal(atTop || onRight ? 0D : 20D, gutter.Left - baseline.Left, 3);
        Assert.Equal(!atTop && onRight ? -20D : 0D, gutter.Right - baseline.Right, 3);
        Assert.Equal(atTop ? -20D : 0D, gutter.Top - baseline.Top, 3);
    }

    [Fact]
    public void SaveAsPdf_ExplicitMarginOverrideReplacesTheAuthoredGutter() {
        var options = new WordToPdfOptions { IncludePageNumbers = false, Margins = PdfCore.PageMargins.Uniform(30) };
        var baseline = RenderGutterMarkers(false, false, WordFileFormat.Docx, 0, options);
        var gutter = RenderGutterMarkers(true, true, WordFileFormat.Docx, 400, options);
        Assert.Equal(baseline, gutter);
    }

    [Fact]
    public void SaveAsPdf_TopGutterLeavesBodySpaceWhenHeadersExpand() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
        document.Sections[0].Margins.Gutter = 1440;
        document.Settings.GutterAtTop = true;
        document.AddHeadersAndFooters();
        for (int index = 0; index < 100; index++) {
            document.Header.Default!.AddParagraph("Header " + index);
        }
        document.AddParagraph("BodyMarker");

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = pdf.GetPage(1);
        Assert.Contains("BodyMarker", page.Text);
        Assert.True(FindWordStartY(page, "BodyMarker") > document.Sections[0].Margins.Bottom / 20D);
    }

    private static (double Left, double Right, double Top) RenderGutterMarkers(bool atTop, bool onRight,
        WordFileFormat format, uint gutter, WordToPdfOptions? options = null) {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.Margins.Left = 800; section.Margins.Right = 1400;
        section.Margins.Top = 1100; section.Margins.Bottom = 1300;
        section.Margins.Gutter = gutter;
        document.Settings.GutterAtTop = atTop; section.RtlGutter = onRight;
        foreach (string marker in new[] { "LeftMarker", "RightMarker" }) {
            var paragraph = document.AddParagraph(marker);
            paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
            paragraph.LineSpacingAfterPoints = 0;
            if (marker == "RightMarker") paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
        }
        using var source = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(source);
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdfBytes(options ?? new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        return (FindWordStartX(page, "LeftMarker"), FindWordStartX(page, "RightMarker"), FindWordStartY(page, "LeftMarker"));
    }
}
