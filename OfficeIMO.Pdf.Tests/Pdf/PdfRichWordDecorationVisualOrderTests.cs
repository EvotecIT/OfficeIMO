using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRichWordDecorationVisualOrderTests {
    [Theory]
    [InlineData("Body")]
    [InlineData("Header")]
    [InlineData("Footer")]
    public void NaturalRichRtlWordUnderlinesFollowUnequalVisualWordsAcrossANonbreakingSpace(string owner) {
        const string text = "\u05D0\u05D1\u00A0\u05D2\u05D3\u05D5\u05D6";
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            Enumerable.Range(32, 95).Concat(new[] { 0x00A0, 0x05D0, 0x05D1, 0x05D2,
                0x05D3, 0x05D5, 0x05D6 }).ToArray());
        var options = new PdfCreateOptions {
            PageWidth = 300, PageHeight = 220, MarginLeft = 30, MarginRight = 30,
            MarginTop = 30, MarginBottom = 30, DefaultFontSize = 12,
            TextShapingMode = PdfTextShapingMode.UnicodeScalar, CompressContentStreams = false
        }.EmbedStandardFont(PdfStandardFont.Helvetica, font, "RTL word decoration");
        var run = new PdfTextRun(text, fontSize: 12, underlineStyle: OfficeTextDecorationStyle.Words);
        PdfDocument document = PdfDocument.Create(options).Compose(composition => composition.Page(page => {
            page.Content(content => {
                if (owner == "Body") content.Paragraph(paragraph => paragraph.Runs(new[] { run }));
                else content.Paragraph(paragraph => paragraph.Text("\u05D0"));
            });
            if (owner == "Header") page.Header(header => header.Text(builder => builder.Run(run)));
            else if (owner == "Footer") page.Footer(footer => footer.Text(builder => builder.Run(run)));
        }));

        using var pdf = PdfPigDocument.Open(document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var strokes = pdf.GetPage(1).Paths.Where(path => path.IsStroked)
            .Select(path => path.GetBoundingRectangle()).Where(bounds => bounds.HasValue)
            .Select(bounds => bounds!.Value).OrderBy(bounds => bounds.Left).ToArray();
        Assert.Equal(2, strokes.Length);
        // The synthetic font gives every glyph a 500/1000-em advance. RTL paints
        // the four-letter logical suffix on the left, then NBSP, then two letters.
        Assert.InRange(strokes[0].Width, 23.99D, 24.01D);
        Assert.InRange(strokes[1].Width, 11.99D, 12.01D);
        Assert.InRange(strokes[1].Left - strokes[0].Right, 5.99D, 6.01D);
    }
}
