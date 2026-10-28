using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false, 78, 1100, 400, 700, -250, 50, 1.5D, .4D)]
    [InlineData(3, true, 78, 1100, 400, 700, -250, 50, 1.5D, .4D)]
    [InlineData(4, false, 78, 1100, 400, 700, -250, 50, 1.5D, .4D)]
    [InlineData(4, true, 78, 1100, 400, 700, -250, 50, 1D, .25D)]
    [InlineData(5, true, 78, 1100, 400, 700, -250, 50, 1D, .25D)]
    [InlineData(4, true, 74, 1100, 400, 700, -250, 50, 1D, .25D)]
    [InlineData(4, true, 73, 1100, 400, 700, -250, 50, 1.2D, .2D)]
    [InlineData(4, true, 78, 1100, 400, 0, 0, 0, 1.5D, .4D)]
    [InlineData(0, false, 78, 800, 200, 700, -250, 50, 1.2D, .2D)]
    public void FontLineMetricsSelectCompatibleHeightAndDescent(int version, bool useTypo, int length,
        int windowsAscent, int windowsDescent, int typoAscent, int typoDescent, int typoGap,
        double expectedAdvance, double expectedDescent) {
        byte[] font = CreateSelectionMetricFont(version, useTypo, length,
            windowsAscent, windowsDescent, typoAscent, typoDescent, typoGap);
        var metrics = OfficeOpenTypeLineMetrics.TryRead(font);
        Assert.True(metrics.HasValue);
        Assert.Equal(expectedAdvance, metrics.Value.HorizontalAdvanceRatio, 6);
        Assert.Equal(expectedDescent, metrics.Value.DescentRatio, 6);
    }

    [Theory]
    [InlineData(false, "body", false)]
    [InlineData(true, "body", false)]
    [InlineData(false, "columns", false)]
    [InlineData(true, "columns", false)]
    [InlineData(false, "table", false)]
    [InlineData(true, "table", false)]
    [InlineData(false, "body", true)]
    [InlineData(true, "body", true)]
    [InlineData(false, "columns", true)]
    [InlineData(true, "columns", true)]
    [InlineData(false, "table", true)]
    [InlineData(true, "table", true)]
    public void SaveAsPdf_FontMetricSelectionControlsNaturalSpacingAndBaseline(bool nativeDoc, string frame, bool useTypo) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = frame == "table"
            ? source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : source.AddParagraph();
        if (frame == "columns") source.Sections[0].ColumnCount = 2;
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(TextLine("A", 12, true), TextLine("B", 12, true), TextLine("C", 12, false));
        paragraph.LineSpacing = 276;
        paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }) : source.ToBytes()));
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            CreateSelectionMetricFont(4, useTypo, 78, 1100, 400, 700, -250, 50), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options,
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(300)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Single(pdf.GetPages());
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var middle = Assert.Single(letters, letter => letter.Value == "B");
        var last = Assert.Single(letters, letter => letter.Value == "C");
        double advance = useTypo ? 13.8D : 20.7D;
        if (frame != "table") Assert.Equal(270D - (useTypo ? 9D : 13.2D), first.StartBaseLine.Y, 3);
        Assert.Equal(advance, first.StartBaseLine.Y - middle.StartBaseLine.Y, 3);
        Assert.Equal(advance, middle.StartBaseLine.Y - last.StartBaseLine.Y, 3);
        Assert.Equal(new[] { 12D, 12D, 12D }, new[] { first.PointSize, middle.PointSize, last.PointSize });
    }

    private static byte[] CreateSelectionMetricFont(int version, bool useTypo, int length,
        int windowsAscent, int windowsDescent, int typoAscent, int typoDescent, int typoGap) {
        byte[] font = CreateBaselineMetricFont();
        int record = FindBaselineFontTableRecord(font, "OS/2");
        int offset = checked((int)(((uint)font[record + 8] << 24) | ((uint)font[record + 9] << 16)
            | ((uint)font[record + 10] << 8) | font[record + 11]));
        WriteBaselineFontUInt16(font, offset, version);
        WriteBaselineFontUInt16(font, offset + 62, useTypo ? 128 : 0);
        WriteBaselineFontUInt16(font, offset + 68, typoAscent);
        WriteBaselineFontUInt16(font, offset + 70, typoDescent);
        WriteBaselineFontUInt16(font, offset + 72, typoGap);
        WriteBaselineFontUInt16(font, offset + 74, windowsAscent);
        WriteBaselineFontUInt16(font, offset + 76, windowsDescent);
        WriteBaselineFontUInt16(font, record + 12, 0);
        WriteBaselineFontUInt16(font, record + 14, length);
        return font;
    }
}
