using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Drawing;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests;

public class HtmlTextExportConsistencyTests {
    private readonly ITestOutputHelper _log;
    public HtmlTextExportConsistencyTests(ITestOutputHelper log) => _log = log;

    [Theory]
    [InlineData("sup")]
    [InlineData("sub")]
    public void OutlinedScriptsStyleTheScaledGlyphWithoutMovingItsItalicOrStrikeGeometry(string script) {
        static OfficeRasterImage Render(string body, double size) {
            string font = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "RobotoFlex.ttf")));
            string html = "<style>@page{size:180px 100px;margin:0}body{margin:0}" +
                "@font-face{font-family:ScriptProof;src:url(data:font/ttf;base64," + font + ")}p{margin:0;font:" +
                size.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px/32px ScriptProof}" +
                "span{font-style:italic;color:blue;text-decoration:line-through;text-decoration-color:red}</style><p>" + body + "</p>";
            var converted = HtmlConversionDocument.Parse(html).ToPdfDocumentResult();
            Assert.Contains(converted.Report.Warnings, warning => warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined);
            return OfficeDrawingRasterRenderer.Render(OfficeIMO.Pdf.PdfDocument.Load(converted.ToBytes()).Render.Drawing(1), 4D, OfficeColor.White);
        }
        OfficeRasterImage reference = Render("<span>H</span>", 24D * 0.65D);
        OfficeRasterImage scripted = Render("<" + script + "><span>H</span></" + script + ">", 24D);
        static int LeftBlue(OfficeRasterImage image) {
            for (int x = 0; x < image.Width; x++)
                for (int y = 0; y < image.Height; y++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (color.B > 160 && color.R < 80 && color.G < 80) return x;
                }
            return -1;
        }
        Assert.InRange(LeftBlue(scripted) - LeftBlue(reference), -1, 1);
        var referenceInk = FindInkBottoms(reference);
        var scriptedInk = FindInkBottoms(scripted);
        Assert.True(referenceInk.Red >= 0 && scriptedInk.Red >= 0);
        Assert.InRange((scriptedInk.Red - scriptedInk.Blue) - (referenceInk.Red - referenceInk.Blue), -1, 1);
    }

    [Fact]
    public void RasterLetterSpacingMovesTheNextGlyphWithoutStretchingEitherGlyph() {
        static (int RedWidth, int BlueLeft, int BlueWidth) Render(double spacing) {
            string html = "<p style='margin:0;font:24px Arial;letter-spacing:" + spacing + "px'>" +
                "<span style='color:red'>H</span><span style='color:blue'>H</span></p>";
            HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html,
                new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 140D, ViewportHeight = 90D, Margins = HtmlRenderMargins.All(0) });
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(scene.Pages[0].CreateDrawing(), 4D, OfficeColor.White);
            int redLeft = image.Width, redRight = -1, blueLeft = image.Width, blueRight = -1;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (color.R > 160 && color.B < 80 && color.G < 80) {
                        redLeft = Math.Min(redLeft, x); redRight = Math.Max(redRight, x);
                    }
                    if (color.B > 160 && color.R < 80 && color.G < 80) {
                        blueLeft = Math.Min(blueLeft, x); blueRight = Math.Max(blueRight, x);
                    }
                }
            }
            Assert.True(redRight >= redLeft && blueRight >= blueLeft);
            return (redRight - redLeft, blueLeft, blueRight - blueLeft);
        }
        var normal = Render(0D);
        var spaced = Render(8D);
        Assert.InRange(spaced.RedWidth - normal.RedWidth, -1, 1);
        Assert.InRange(spaced.BlueWidth - normal.BlueWidth, -1, 1);
        Assert.InRange(spaced.BlueLeft - normal.BlueLeft, 31, 33);
    }

    [Theory]
    [InlineData(false, "sup")]
    [InlineData(false, "sub")]
    [InlineData(true, "sup")]
    [InlineData(true, "sub")]
    public void OutlinedPdfKeepsScriptDisplacementRelativeToTheUnscaledBaseline(bool foreignObject, string script) {
        string font = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "RobotoFlex.ttf")));
        string styles = "<style>@page{size:180px 100px;margin:0}body{margin:0}" +
            "@font-face{font-family:ScriptProof;src:url(data:font/ttf;base64," + font + ")}p{margin:0;font:24px/32px ScriptProof}</style>";
        string body = "<p><span style='color:blue'>H</span><" + script + " style='color:red'>H</" + script + "></p>";
        HtmlRenderDocument reference = HtmlRenderTestDriver.Render(styles + body);
        double expectedOffset = Assert.Single(reference.Pages[0].Visuals.OfType<HtmlRenderText>(), run => run.Baseline != OfficeTextBaseline.Normal).BaselineOffset;
        if (foreignObject) body = "<svg xmlns='http://www.w3.org/2000/svg' width='180' height='100'><foreignObject width='180' height='100'><div xmlns='http://www.w3.org/1999/xhtml'>" + body + "</div></foreignObject></svg>";
        var converted = HtmlConversionDocument.Parse(styles + body).ToPdfDocumentResult();
        Assert.Contains(converted.Report.Warnings, warning => warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined);
        OfficeDrawing drawing = OfficeIMO.Pdf.PdfDocument.Load(converted.ToBytes()).Render.Drawing(1);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, scale: 4D, background: OfficeColor.White);
        (int blueBottom, int redBottom) = FindInkBottoms(image);
        Assert.True(blueBottom >= 0 && redBottom >= 0);
        Assert.InRange((redBottom - blueBottom) / 3D, expectedOffset - 0.5D, expectedOffset + 0.5D);
    }

    [Theory]
    [InlineData("sup")]
    [InlineData("sub")]
    public void DirectRasterKeepsScriptDisplacementRelativeToTheUnscaledBaseline(string script) {
        string html = "<p style='margin:0;font:24px/32px Arial'><span style='color:blue'>H</span>" +
            "<" + script + " style='color:red'>H</" + script + "></p>";
        HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 140D, ViewportHeight = 90D });
        HtmlRenderText scripted = Assert.Single(scene.Pages[0].Visuals.OfType<HtmlRenderText>(), run => run.Baseline != OfficeTextBaseline.Normal);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(scene.Pages[0].CreateDrawing(), scale: 4D, background: OfficeColor.White);
        (int blueBottom, int redBottom) = FindInkBottoms(image);
        Assert.True(blueBottom >= 0 && redBottom >= 0);
        Assert.InRange((redBottom - blueBottom) / 4D, scripted.BaselineOffset - 0.5D, scripted.BaselineOffset + 0.5D);
    }

    private static (int Blue, int Red) FindInkBottoms(OfficeRasterImage image) {
        int blueBottom = -1, redBottom = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.B > 160 && color.R < 80 && color.G < 80) blueBottom = y;
                if (color.R > 160 && color.B < 80 && color.G < 80) redBottom = y;
            }
        }
        return (blueBottom, redBottom);
    }

    [Theory]
    [InlineData("plain", "<text x='10' y='35' font-family='Arial' font-size='24'>NORTHWIND</text>", "NORTHWIND")]
    [InlineData("preserved-spaces", "<text x='10' y='35' font-family='Arial' font-size='24' xml:space='preserve'>A  B</text>", "A  B")]
    [InlineData("text-length", "<text x='10' y='35' font-family='Arial' font-size='24' textLength='80' lengthAdjust='spacingAndGlyphs'>NORTHWIND</text>", "NORTHWIND")]
    [InlineData("spaced", "<text x='10' y='35' font-family='Arial' font-size='24' letter-spacing='4'>NORTHWIND</text>", "NORTHWIND")]
    [InlineData("tspan", "<text x='10' y='35' font-family='Arial' font-size='24'>A <tspan font-weight='bold'>B</tspan> C</text>", "A B C")]
    public void SvgTextKeepsItsContent(string name, string inner, string expected) {
        string html = "<svg xmlns='http://www.w3.org/2000/svg' width='320' height='80' viewBox='0 0 320 80'>" + inner + "</svg>";
        var source = HtmlConversionDocument.Parse(html);
        var scene = HtmlRenderTestDriver.Render(source);
        foreach (var drawing in scene.Pages[0].Visuals.OfType<HtmlRenderDrawing>())
            foreach (var run in drawing.Drawing.Elements.OfType<OfficeDrawingText>()) _log.WriteLine($"Scene [{run.Text}] x={run.X} width={run.Width} advance={run.TextAdvanceWidth}");
        byte[] bytes = source.ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.ToArray();
        string text = string.Concat(letters.Select(x => x.Value));
        _log.WriteLine(name + " text=" + text);
        foreach (var letter in letters) _log.WriteLine($"{letter.Value}: {letter.StartBaseLine.X:F3}..{letter.EndBaseLine.X:F3} y={letter.StartBaseLine.Y:F3}");
        Assert.Equal(expected, text);
        Assert.Single(letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct());
        if (name == "text-length") Assert.InRange(letters.Last().EndBaseLine.X-letters.First().StartBaseLine.X, 59.9, 60.1);
    }

    [Fact]
    public void LetterSpacingKeepsGlyphPaintWidthSeparateFromAdvance() {
        const string html = "<p style='font-family:Arial;font-size:24px;letter-spacing:8px'>AB</p>";
        var source = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source);
        HtmlRenderText[] runs = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        foreach (var run in runs) _log.WriteLine($"{run.Text}: paint={run.TextPaintWidth} advance={run.TextAdvanceWidth}");
        string svg = source.ToSvg();
        var xml = System.Xml.Linq.XDocument.Parse(svg);
        var textNodes = xml.Descendants().Where(element => element.Name.LocalName == "text").ToArray();
        foreach (var text in textNodes) _log.WriteLine(text.ToString());
        Assert.NotEmpty(runs);
        Assert.Equal(runs.Length, textNodes.Length);
        for (int index=0;index<runs.Length;index++) {
            double writtenWidth = double.Parse(textNodes[index].Attribute("textLength")!.Value, System.Globalization.CultureInfo.InvariantCulture);
            Assert.InRange(writtenWidth, runs[index].TextPaintWidth!.Value-0.1, runs[index].TextPaintWidth!.Value+0.1);
        }
    }

    [Theory]
    [InlineData(false, "sup")]
    [InlineData(true, "sup")]
    [InlineData(false, "sub")]
    [InlineData(true, "sub")]
    public void SvgForeignObjectKeepsScripts(bool nested, string script) {
        string html = "<p style='margin:0;font-family:Arial;font-size:24px'>A<"+script+">2</"+script+">B</p>";
        if (nested) html = "<svg xmlns='http://www.w3.org/2000/svg' width='320' height='100'><foreignObject width='320' height='100'><div xmlns='http://www.w3.org/1999/xhtml'>"+html+"</div></foreignObject></svg>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("A2B", string.Concat(letters.Select(letter => letter.Value)));
        foreach (var letter in letters) _log.WriteLine($"{letter.Value}: size={letter.FontSize} y={letter.StartBaseLine.Y}");
        Assert.True(letters[1].FontSize < letters[0].FontSize * 0.9);
        Assert.True(script == "sup" ? letters[1].StartBaseLine.Y > letters[0].StartBaseLine.Y : letters[1].StartBaseLine.Y < letters[0].StartBaseLine.Y);
    }
}
