using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "body")]
    [InlineData(true, "body")]
    [InlineData(false, "columns")]
    [InlineData(true, "columns")]
    [InlineData(false, "table")]
    [InlineData(true, "table")]
    public void SaveAsPdf_EmptyLeadingRunsDoNotEnlargeVisibleFontLineSpacing(bool nativeDoc, string frame) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = frame == "table"
            ? source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : source.AddParagraph();
        if (frame == "columns") source.Sections[0].ColumnCount = 2;
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(new Run(new RunProperties(new RunFonts { Ascii = "TallFallback", HighAnsi = "TallFallback" })),
            TextLine("A", 24, true), TextLine("B", 24, true), TextLine("C", 24, false));
        paragraph.LineSpacing = 240; paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        var options = new PdfOptions();
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("TallFallback",
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1200, -300, 100, 300, Enumerable.Range(32, 95).ToArray())));
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Arial", CreateBaselineMetricFont()));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options,
            PageSize = new OfficeIMO.Pdf.PageSize(300, 300), Margins = PageMargins.Uniform(30)
        }));
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var second = Assert.Single(letters, letter => letter.Value == "B");
        var third = Assert.Single(letters, letter => letter.Value == "C");
        Assert.Equal(28.8D, first.StartBaseLine.Y - second.StartBaseLine.Y, 3);
        Assert.Equal(28.8D, second.StartBaseLine.Y - third.StartBaseLine.Y, 3);
        Assert.Equal(new[] { 24D, 24D, 24D }, new[] { first.PointSize, second.PointSize, third.PointSize });
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
    public void SaveAsPdf_MixedSizeBaselinesUseFontLineBoxes(bool nativeDoc, string frame, bool minimum) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = frame == "table"
            ? source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : source.AddParagraph();
        if (frame == "columns") source.Sections[0].ColumnCount = 2;
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(TextLine("A", 8, true), TextLine("B", 32, true), TextLine("C", 8, false));
        paragraph.LineSpacing = minimum ? 400 : 240;
        paragraph.LineSpacingRule = minimum ? WordLineSpacingRule.AtLeast : WordLineSpacingRule.Auto;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica, CreateBaselineMetricFont(), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options,
            PageSize = new OfficeIMO.Pdf.PageSize(300, 300), Margins = PageMargins.Uniform(30)
        }));
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var middle = Assert.Single(letters, letter => letter.Value == "B");
        var last = Assert.Single(letters, letter => letter.Value == "C");
        // 1.2-em natural advance and .3-em Windows descent position the
        // baseline .9 em below the line top; minimum spacing adds room above.
        if (frame != "table") Assert.Equal(270D - (minimum ? 17.6D : 7.2D), first.StartBaseLine.Y, 3);
        Assert.Equal(31.2D, first.StartBaseLine.Y - middle.StartBaseLine.Y, 3);
        Assert.Equal(minimum ? 27.2D : 16.8D, middle.StartBaseLine.Y - last.StartBaseLine.Y, 3);
        Assert.Equal(new[] { 8D, 32D, 8D }, new[] { first.PointSize, middle.PointSize, last.PointSize });
    }

    [Theory]
    [InlineData(68, .2D)]
    [InlineData(78, .3D)]
    public void FontLineMetricsRespectTheOs2TableBoundary(int os2Length, double expectedDescent) {
        byte[] font = CreateBaselineMetricFont();
        int record = FindBaselineFontTableRecord(font, "OS/2");
        WriteBaselineFontUInt16(font, record + 12, 0);
        WriteBaselineFontUInt16(font, record + 14, os2Length);
        var metrics = OfficeOpenTypeLineMetrics.TryRead(font);
        Assert.True(metrics.HasValue);
        Assert.Equal(1.2D, metrics.Value.HorizontalAdvanceRatio, 6);
        Assert.Equal(expectedDescent, metrics.Value.WindowsDescentRatio, 6);
    }

    private static Run TextLine(string text, int size, bool lineBreak) {
        var run = new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
            new FontSize { Val = (size * 2).ToString(System.Globalization.CultureInfo.InvariantCulture) }), new Text(text));
        if (lineBreak) run.Append(new Break());
        return run;
    }

    private static byte[] CreateBaselineMetricFont() =>
        ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(900, -200, 100, 300, Enumerable.Range(32, 95).ToArray());

    private static int FindBaselineFontTableRecord(byte[] font, string tag) {
        int count = (font[4] << 8) | font[5];
        for (int i = 0; i < count; i++) {
            int record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(font, record, 4) == tag) return record;
        }
        throw new InvalidOperationException("Missing test font table: " + tag);
    }

    private static void WriteBaselineFontUInt16(byte[] font, int offset, int value) {
        font[offset] = (byte)(value >> 8); font[offset + 1] = (byte)value;
    }
}
