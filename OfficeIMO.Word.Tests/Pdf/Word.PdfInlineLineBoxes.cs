using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    // Body Face at 10 pt has a 13 pt font line box (8 pt ascent, 5 pt descent), so a shifted
    // run would reach past it; multiple 1.2 spacing gives a 15.6 pt pitch with its leading
    // below the baseline.
    private const double InlineLineBoxPitch = 15.6D;

    // Word sizes superscript and subscript runs without their shift, so they never change the line pitch.
    [Theory]
    [InlineData(WordVerticalTextPosition.Superscript)]
    [InlineData(WordVerticalTextPosition.Subscript)]
    public void SaveAsPdf_ShiftedRunKeepsTheParagraphLinePitch(WordVerticalTextPosition position) {
        using WordDocument document = CreateInlineLineBoxDocument(out WordParagraph paragraph);
        paragraph.AddText("Beta x");
        paragraph.AddText("2").SetVerticalTextAlignment(position);
        paragraph.AddText(" gamma");
        paragraph.AddBreak();
        paragraph.AddText("Delta");

        double[] baselines = ReadInlineLineBoxBaselines(document);

        Assert.Equal(InlineLineBoxPitch, baselines[0] - baselines[1], 3);
        Assert.Equal(InlineLineBoxPitch, baselines[1] - baselines[2], 3);
    }

    // Word sizes and places superscript and subscript text from the font's OS/2 script metrics: the
    // size rounded to half a point, the offset limited to 35% (super) or 8% (sub) of the run size.
    [Theory]
    [InlineData(WordVerticalTextPosition.Superscript, 6D, 3D)]
    [InlineData(WordVerticalTextPosition.Subscript, 6D, -0.5D)]
    public void SaveAsPdf_ShiftedRunTakesItsSizeAndOffsetFromTheFont(WordVerticalTextPosition position, double size, double offset) {
        using WordDocument document = CreateInlineLineBoxDocument(out WordParagraph paragraph);
        paragraph.AddText("Beta x");
        paragraph.AddText("2").SetVerticalTextAlignment(position);

        // Script size 0.6 em; superscript offset 0.3 em; subscript offset 0.05 em.
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -500, 0, 500, Enumerable.Range(32, 95).ToArray());
        ManagedTextShapingTestAssets.WithScriptMetrics(font, subscriptSize: 600, subscriptOffset: 50, superscriptSize: 600, superscriptOffset: 300);
        using var pdf = PdfPigDocument.Open(RenderInlineLineBoxes(document, font));

        var letters = pdf.GetPage(1).Letters;
        UglyToad.PdfPig.Content.Letter shifted = letters.Single(letter => letter.Value == "2");
        UglyToad.PdfPig.Content.Letter body = letters.First(letter => letter.Value == "x");
        Assert.Equal(size, shifted.PointSize, 3);
        Assert.Equal(offset, shifted.StartBaseLine.Y - body.StartBaseLine.Y, 3);
    }

    // A taller inline picture raises its line; multiple spacing keeps its leading below the baseline.
    // Word also lays the picture out with the bottom effect extent it records, here 0.5 pt.
    [Theory]
    [InlineData(0L, 47.6D)]
    [InlineData(6350L, 48.1D)]
    public void SaveAsPdf_InlinePictureUnderMultipleSpacingKeepsTheNextLinePitch(long bottomExtent, double pictureLineGap) {
        using WordDocument document = CreateInlineLineBoxDocument(out WordParagraph paragraph);
        paragraph.AddText("Beta ");
        byte[] png = OfficeIMO.Drawing.OfficeRasterImageEncoder.Encode(
            new OfficeIMO.Drawing.OfficeRasterImage(2, 2, OfficeIMO.OfficeColor.Black), OfficeIMO.Drawing.OfficeImageExportFormat.Png);
        WordImage picture;
        using (var stream = new MemoryStream(png))
            picture = paragraph.AddText(string.Empty).AddImage(stream, "picture.png", OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Points(40));
        picture._Image.Inline!.EffectExtent!.BottomEdge = bottomExtent;
        paragraph.AddText(" gamma");
        paragraph.AddBreak();
        paragraph.AddText("Delta");

        double[] baselines = ReadInlineLineBoxBaselines(document);

        // 7.6 pt of the first line below its baseline, the extent, then the 40 pt picture.
        Assert.Equal(pictureLineGap, baselines[0] - baselines[1], 3);
        Assert.Equal(InlineLineBoxPitch, baselines[1] - baselines[2], 3);
    }

    private static WordDocument CreateInlineLineBoxDocument(out WordParagraph paragraph) {
        WordDocument document = WordDocument.Create();
        paragraph = document.AddParagraph("Alpha");
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple(1.2D);
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph.AddBreak();
        return document;
    }

    private static double[] ReadInlineLineBoxBaselines(WordDocument document) {
        using var pdf = PdfPigDocument.Open(RenderInlineLineBoxes(document,
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -500, 0, 500, Enumerable.Range(32, 95).ToArray())));
        double[] baselines = pdf.GetPage(1).GetWords().Where(word => word.Text is "Alpha" or "Beta" or "Delta")
            .Select(word => word.Letters[0].StartBaseLine.Y).OrderByDescending(y => y).ToArray();
        Assert.Equal(3, baselines.Length);
        return baselines;
    }

    // Every run, including breaks and the paragraph mark, uses Body Face at 10 pt.
    private static byte[] RenderInlineLineBoxes(WordDocument document, byte[] font) {
        var body = document._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        foreach (var run in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
            (run.RunProperties ??= new DocumentFormat.OpenXml.Wordprocessing.RunProperties()).Append(InlineLineBoxFont(), InlineLineBoxSize());
        foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
            (paragraph.ParagraphProperties!.ParagraphMarkRunProperties ??= new DocumentFormat.OpenXml.Wordprocessing.ParagraphMarkRunProperties())
                .Append(InlineLineBoxFont(), InlineLineBoxSize());
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Body Face", font));
        return document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }).ToBytes(new PdfSaveOptions { AllowLoss = true });
    }

    private static DocumentFormat.OpenXml.Wordprocessing.RunFonts InlineLineBoxFont() =>
        new() { Ascii = "Body Face", HighAnsi = "Body Face" };

    private static DocumentFormat.OpenXml.Wordprocessing.FontSize InlineLineBoxSize() => new() { Val = "20" };
}
