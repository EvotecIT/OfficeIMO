using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlArabicProvider_DefaultLinkUnderlinePreservesContextGlyphsInRasterAndSvg() {
        byte[] font = PreparedArabicProviderFont();
        const string href = "https://example.com/arabic";
        // An ordinary taller line leaves room to inspect the default underline
        // independently of the glyph ink and the continuous surface boundary.
        const string css = "#probe{line-height:40px}";
        HtmlRenderDocument linked = RenderArabicContext("<a href='" + href + "'>" + PreparedArabicColorWord + "</a>",
            extraCss: css, provider: new PreparedArabicContextProvider(font), fontData: font);
        HtmlRenderDocument plain = RenderArabicContext("<a href='" + href + "' style='text-decoration:none'>"
            + PreparedArabicColorWord + "</a>", extraCss: css,
            provider: new PreparedArabicContextProvider(font), fontData: font);
        HtmlRenderText[] text = ArabicContextText(linked).ToArray();
        HtmlRenderShape[] bands = DecorationLines(linked)
            .Where(shape => shape.Source!.EndsWith(":underline", StringComparison.Ordinal)).ToArray();

        Assert.Equal(ArabicContextWord, linked.Text.Trim());
        Assert.Equal(plain.Text, linked.Text);
        Assert.All(text, run => {
            Assert.NotNull(run.PreparedShaping);
            Assert.Equal(href, run.LinkUri);
            Assert.Equal(OfficeTextDecorationStyle.None, run.UnderlineStyle);
        });
        Assert.Equal(2, bands.Length);
        Assert.Equal(20D, bands.Sum(band => band.Width), 6);
        Assert.All(bands, band => Assert.Equal(OfficeColor.FromRgb(0, 0, 238), band.Shape.FillColor));
        Assert.Empty(DecorationLines(plain));
        OfficeDrawing drawing = Assert.Single(linked.Pages).CreateDrawing();
        OfficeDrawing undecorated = Assert.Single(plain.Pages).CreateDrawing();
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(
            OfficeDrawingSvgExporter.ToSvg(drawing)), out OfficeDrawing? svgDrawing));
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(
            OfficeDrawingSvgExporter.ToSvg(undecorated)), out OfficeDrawing? svgUndecorated));
        var pairs = new[] {
            (Backend: "Drawing", Decorated: OfficeDrawingRasterRenderer.Render(drawing), Plain: OfficeDrawingRasterRenderer.Render(undecorated), ScaleX: 1D, ScaleY: 1D),
            (Backend: "SVG", Decorated: OfficeDrawingRasterRenderer.Render(svgDrawing!),
                Plain: OfficeDrawingRasterRenderer.Render(svgUndecorated!),
                ScaleX: svgDrawing!.Width / drawing.Width, ScaleY: svgDrawing.Height / drawing.Height)
        };
        foreach (var pair in pairs) {
            Assert.Equal(pair.Plain.Width, pair.Decorated.Width);
            Assert.Equal(pair.Plain.Height, pair.Decorated.Height);
            // Removing the link decoration must leave the joined colored glyphs
            // byte-for-byte unchanged above the decoration bands in each backend.
            // SVG declares physical units. Use the parsed surface's exact scale,
            // before raster dimensions round up to whole device pixels.
            int glyphBottom = Math.Max(0, (int)Math.Floor(bands.Min(band => band.Y) * pair.ScaleY) - 1);
            bool redInk = false, blueInk = false;
            for (int y = 0; y < Math.Min(glyphBottom, pair.Decorated.Height); y++) {
                for (int x = 0; x < pair.Decorated.Width; x++) {
                    OfficeColor pixel = pair.Decorated.GetPixel(x, y);
                    Assert.Equal(pair.Plain.GetPixel(x, y), pixel);
                    redInk |= pixel.R > 200 && pixel.G < 80 && pixel.B < 80;
                    blueInk |= pixel.B > 200 && pixel.R < 80 && pixel.G < 80;
                }
            }
            Assert.True(redInk && blueInk, "Both colored context fragments must retain actual glyph ink.");
            foreach (HtmlRenderShape band in bands) {
                int left = Math.Max(0, (int)Math.Floor(band.X * pair.ScaleX));
                int right = Math.Min(pair.Decorated.Width, (int)Math.Ceiling((band.X + band.Width) * pair.ScaleX));
                int top = Math.Max(0, (int)Math.Floor(band.Y * pair.ScaleY));
                int bottom = Math.Min(pair.Decorated.Height, (int)Math.Ceiling((band.Y + band.Height) * pair.ScaleY));
                bool visible = Enumerable.Range(top, Math.Max(0, bottom - top)).Any(y =>
                    Enumerable.Range(left, Math.Max(0, right - left)).Any(x =>
                        pair.Decorated.GetPixel(x, y) != pair.Plain.GetPixel(x, y)));
                if (!visible) {
                    var changed = Enumerable.Range(0, pair.Decorated.Width * pair.Decorated.Height)
                        .Where(index => pair.Decorated.GetPixel(index % pair.Decorated.Width, index / pair.Decorated.Width)
                            != pair.Plain.GetPixel(index % pair.Plain.Width, index / pair.Plain.Width)).ToArray();
                    Assert.Fail($"{pair.Backend} {pair.Decorated.Width}x{pair.Decorated.Height}: underline [{left},{top}..{right},{bottom}] has no visible paint; "
                        + $"changed pixels={changed.Length}, rows=" + string.Join(",", changed.Select(index => index / pair.Decorated.Width).Distinct()));
                }
            }
        }
        linked.RequireNoLoss();
    }
}
