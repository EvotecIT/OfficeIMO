using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingParagraphOverflowTests {
    private const string Caption = "UnwrappedOverflowCaption";
    private const string Link = "https://example.test/caption";

    [Theory]
    [InlineData(OfficeTextAlignment.Center, false)]
    [InlineData(OfficeTextAlignment.Right, false)]
    [InlineData(OfficeTextAlignment.Center, true)]
    [InlineData(OfficeTextAlignment.Right, true)]
    public void UnwrappedParagraphLinkCoversItsCompleteOverwideText(OfficeTextAlignment alignment, bool embedded) {
        byte[] bytes = CreateCaption(alignment, embedded, wrap: false);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(Caption, string.Concat(letters.Select(letter => letter.Value)));
        double left = letters.Min(letter => letter.StartBaseLine.X);
        double right = letters.Max(letter => letter.EndBaseLine.X);
        Assert.True(left < 130D);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(Link));
        Assert.InRange(Math.Abs(link.X1 - left), 0D, .02D);
        Assert.InRange(Math.Abs(link.X2 - right), 0D, .02D);
        Assert.InRange(link.Y1, 40D, 80D);
        Assert.InRange(link.Y2, 40D, 80D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedParagraphLinksRemainInsideTheContentFrame(bool embedded) {
        byte[] bytes = CreateCaption(OfficeTextAlignment.Center, embedded, wrap: true);
        var links = PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri(Link).ToArray();
        Assert.NotEmpty(links);
        Assert.All(links, link => {
            Assert.InRange(link.X1, 130D, 170D);
            Assert.InRange(link.X2, 130D, 170D);
            Assert.InRange(link.Y1, 40D, 80D);
            Assert.InRange(link.Y2, 40D, 80D);
        });
    }

    [Fact]
    public void UnwrappedSyntheticItalicInkFitsTheHorizontalClipWithoutExpandingItsHeight() {
        const double size = 20D;
        byte[] fontBytes = ManagedTextShapingTestAssets.CreateFont('A');
        var drawing = new OfficeDrawing(200D, 100D);
        drawing.Fonts.Add("Ink Probe", fontBytes);
        drawing.AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A", size, OfficeColor.Black,
                italic: true, fontFamily: "Ink Probe") }, OfficeTextAlignment.Center)
        }, 100D, 20D, 4D, 40D, wrapText: false);
        byte[] bytes = Render(drawing, embedded: true);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        Assert.Equal("A", letter.Value);
        // This minimal fixture omits the head magic number, so PdfPig substitutes
        // an em box. Read the actual rectangle glyph directly from its SFNT tables.
        var font = new UglyToad.PdfPig.Fonts.TrueType.TrueTypeDataBytes(fontBytes);
        font.Seek(FindTable("head") + 18);
        int unitsPerEm = font.ReadUnsignedShort();
        font.Seek(FindTable("head") + 50);
        Assert.Equal(0, font.ReadSignedShort()); // Short loca offsets.
        font.Seek(FindTable("loca") + 2); // CreateFont maps A to glyph 1.
        int glyphOffset = font.ReadUnsignedShort() * 2;
        font.Seek(FindTable("glyf") + glyphOffset);
        Assert.Equal(1, font.ReadSignedShort());
        double glyphLeft = font.ReadSignedShort(), glyphBottom = font.ReadSignedShort();
        double glyphRight = font.ReadSignedShort(), glyphTop = font.ReadSignedShort();
        Assert.Equal((0D, 0D, 400D, 700D), (glyphLeft, glyphBottom, glyphRight, glyphTop));
        const string number = @"(-?\d+(?:\.\d+)?)";
        string source = Encoding.ASCII.GetString(bytes);
        Match matrix = Regex.Matches(source,
            number + @"\s+" + number + @"\s+" + number + @"\s+" + number + @"\s+" +
            number + @"\s+" + number + @"\s+Tm").Cast<Match>()
            .First(match => Parse(match, 3) != 0D);
        Assert.InRange(Parse(matrix, 3), .333D, .334D);
        double scale = size / unitsPerEm;
        var ink = new[] { (X: glyphLeft, Y: glyphBottom), (X: glyphRight, Y: glyphBottom),
            (X: glyphLeft, Y: glyphTop), (X: glyphRight, Y: glyphTop) }
            .Select(point => (X: scale * (Parse(matrix, 1) * point.X + Parse(matrix, 3) * point.Y) + Parse(matrix, 5),
                Y: scale * (Parse(matrix, 2) * point.X + Parse(matrix, 4) * point.Y) + Parse(matrix, 6))).ToArray();
        Assert.True(ink.Max(point => point.X) > letter.EndBaseLine.X + 1D);
        Match clip = Assert.Single(Regex.Matches(source,
            number + @"\s+" + number + @"\s+" + number + @"\s+" + number + @"\s+re\s+W\s+n")
            .Cast<Match>());
        double x = Parse(clip, 1), y = Parse(clip, 2), width = Parse(clip, 3), height = Parse(clip, 4);
        Assert.True(x <= ink.Min(point => point.X));
        Assert.True(x + width >= ink.Max(point => point.X));
        Assert.Equal(40D, y);
        Assert.Equal(40D, height);

        long FindTable(string tag) {
            font.Seek(4);
            int count = font.ReadUnsignedShort();
            for (int index = 0; index < count; index++) {
                font.Seek(12 + index * 16);
                string current = font.ReadTag();
                font.ReadUnsignedInt();
                uint offset = font.ReadUnsignedInt();
                if (current == tag) return offset;
            }
            throw new InvalidOperationException($"The ink fixture lacks its {tag} table.");
        }

        static double Parse(Match match, int group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
    }

    private static byte[] CreateCaption(OfficeTextAlignment alignment, bool embedded, bool wrap) {
        var drawing = new OfficeDrawing(300D, 100D);
        string family = embedded ? "Proof Sans" : "Helvetica";
        if (embedded) drawing.Fonts.Add(family, File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!));
        drawing.AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(Caption, 10D, OfficeColor.Black,
                italic: true, underline: true, fontFamily: family) { LinkUri = Link } }, alignment)
        }, 130D, 20D, 40D, 40D, wrapText: wrap);
        return Render(drawing, embedded);
    }

    private static byte[] Render(OfficeDrawing drawing, bool embedded) {
        var options = new PdfOptions { PageWidth = drawing.Width, PageHeight = drawing.Height,
            MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D, CompressContentStreams = false };
        if (embedded) options.UseRenderingProfile(new OfficeRenderingProfile("overflow", drawing.Fonts));
        return PdfDocument.Create(options).Compose(canvas => canvas.Page(page =>
            page.Content(content => content.Drawing(drawing)))).ToBytes();
    }
}
