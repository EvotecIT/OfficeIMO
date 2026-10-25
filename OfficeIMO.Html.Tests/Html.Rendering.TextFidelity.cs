using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlText_MixedInlineEmphasisPreservesAuthoredWordBoundaries() {
        const string html = "<p style='margin:0;font:16px Arial'>Before <strong>bold</strong> after <em>italic</em> tail</p>";

        IReadOnlyList<HtmlRenderText> text = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 320D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        }).Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .OrderBy(visual => visual.PaintOrder)
            .ToList();

        Assert.Equal("Before bold after italic tail", string.Concat(text.Select(visual => visual.Text)));
        Assert.Single(text.Select(visual => Math.Round(visual.Y, 3)).Distinct());
    }

    [Fact]
    public void HtmlImageExport_UsesPixelSvgAndExactSharedTextAdvances() {
        const string html = "<html style='margin:0'><body style='margin:0'>"
            + "<p style='margin:0;font:14px Arial'><strong>One model.</strong> Shared output.</p>"
            + "<p dir='rtl' style='margin:0;font:14px Arial'>שלום 123</p>"
            + "</body></html>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 320D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            Scale = 1.5D
        };

        HtmlRenderDocument rendered = html.RenderHtml(options);
        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        IReadOnlyList<HtmlRenderText> renderedText = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();
        IReadOnlyList<OfficeDrawingText> textElements = drawing.Elements.OfType<OfficeDrawingText>().ToList();
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        string svg = document.ToSvg(options);
        byte[] png = document.ToPng(options);
        XDocument svgDocument = XDocument.Parse(svg);

        Assert.NotEmpty(textElements);
        Assert.Equal(renderedText.Count, textElements.Count);
        for (int index = 0; index < textElements.Count; index++) {
            OfficeDrawingText drawingText = textElements[index];
            HtmlRenderText sourceText = renderedText[index];
            Assert.Equal(OfficeTextOverflowBehavior.Clip, drawingText.OverflowBehavior);
            Assert.Equal(sourceText.Width, drawingText.Width, 6);
            Assert.Equal(
                Assert.IsType<double>(sourceText.TextAdvanceWidth),
                Assert.IsType<double>(drawingText.TextAdvanceWidth),
                6);
        }
        Assert.Equal("480px", svgDocument.Root!.Attribute("width")!.Value);
        Assert.Equal("150px", svgDocument.Root.Attribute("height")!.Value);
        double[] bounds = svgDocument.Root.Attribute("viewBox")!.Value
            .Split(' ').Select(value => double.Parse(value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(4, bounds.Length);
        Assert.Equal(0D, bounds[0]);
        Assert.Equal(0D, bounds[1]);
        Assert.Equal(320D / 100D, bounds[2] / bounds[3], 6);
        Assert.DoesNotContain("pt\"", svg, StringComparison.Ordinal);
        // No caller font is registered for this default-font workflow. The
        // serialized PDF owns its selected fallback program and glyphs.
        byte[] savedPdf = document.ToPdfBytes(new HtmlToPdfOptions(options));
        var savedPage = OfficeIMO.Pdf.PdfReadDocument.Open(savedPdf).Pages[0];
        OfficeImageExportResult savedSvg = savedPage.ExportImage(OfficeImageExportFormat.Svg,
            new OfficeIMO.Pdf.PdfImageExportOptions {
                Scale = options.Scale * HtmlRenderOptions.CssPixelsPerInch / 72D,
                BackgroundColor = options.BackgroundColor
            });
        Assert.Equal(savedSvg.Bytes, Encoding.UTF8.GetBytes(svg));
        string inkSource = string.Concat(svgDocument.Descendants()
            .Where(element => element.Name.LocalName == "path"
                && !string.IsNullOrWhiteSpace(element.Attribute("d")?.Value))
            .Select(element => element.AncestorsAndSelf()
                .Select(ancestor => ancestor.Attribute("aria-label")?.Value)
                .FirstOrDefault(label => label != null)));
        Assert.Contains("One model.", inkSource, StringComparison.Ordinal);
        Assert.Contains("Shared output.", inkSource, StringComparison.Ordinal);

        using var independentPdf = UglyToad.PdfPig.PdfDocument.Open(savedPdf);
        var saved = independentPdf.GetPage(1);
        HtmlRenderText[] firstLineSource = renderedText.Where(text => text.Text.Any(IsLatinLetter)).ToArray();
        Assert.NotEmpty(firstLineSource);
        var latinLetters = saved.Letters.Where(letter => letter.Value.Any(IsLatinLetter)).ToArray();
        Assert.NotEmpty(latinLetters);
        double expectedBaseline = PaintBaseline(firstLineSource[0]);
        Assert.All(firstLineSource, text => Assert.Equal(expectedBaseline, PaintBaseline(text), 6));
        Assert.All(latinLetters, letter => Assert.InRange(
            Math.Abs((saved.Height - letter.StartBaseLine.Y) / 0.75D - expectedBaseline), 0D, 0.02D));
        double savedBaseline = latinLetters[0].StartBaseLine.Y;
        var firstLineLetters = saved.Letters.Where(letter =>
            Math.Abs(letter.StartBaseLine.Y - savedBaseline) < 0.001D).ToArray();
        Assert.InRange(Math.Abs(firstLineLetters.Min(letter => letter.StartBaseLine.X) / 0.75D
            - firstLineSource.Min(text => text.X)), 0D, 0.02D);
        Assert.InRange(Math.Abs(firstLineLetters.Max(letter => letter.EndBaseLine.X) / 0.75D
            - firstLineSource.Max(text => text.X + text.TextAdvanceWidth!.Value)), 0D, 0.02D);

        static bool IsLatinLetter(char scalar) => scalar is >= 'A' and <= 'Z' or >= 'a' and <= 'z';
        Assert.Contains("One model.", SvgPaintedText.Read(svg), StringComparison.Ordinal);
        Assert.Contains("Shared output.", SvgPaintedText.Read(svg), StringComparison.Ordinal);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
    }

    [Fact]
    public void HtmlImageExport_ClipsTheTextFrameWithoutLosingItsResolvedAdvance() {
        const string html = "<html style='margin:0'><body style='margin:0'>"
            + "<div style='width:20px;white-space:nowrap;font-family:Arial;font-size:100px;line-height:100px'>WW</div>"
            + "</body></html>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 20D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = html.RenderHtml(options);
        HtmlRenderText text = rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .First(item => item.TextAdvanceWidth.HasValue && item.TextAdvanceWidth.Value > item.Width);
        double advance = Assert.IsType<double>(text.TextAdvanceWidth);

        Assert.True(advance > text.Width);
        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        OfficeDrawingText drawingText = drawing.Elements
            .OfType<OfficeDrawingText>()
            .First(item => item.Text == text.Text && Math.Abs(item.X - text.X) < 0.000001D && Math.Abs(item.Y - text.Y) < 0.000001D);
        Assert.Equal(text.Width, drawingText.Width, 6);
        Assert.Equal(advance, Assert.IsType<double>(drawingText.TextAdvanceWidth), 6);
        Assert.NotEmpty(HtmlConversionDocument.Parse(html).ToPng(options));
        Assert.StartsWith("<svg", HtmlConversionDocument.Parse(html).ToSvg(options), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlText_BreakSpacesWrapsAndRetainsEveryPreservedSpace() {
        const string html = "<p style='margin:0;width:18px;font:16px Arial;line-height:18px'>A<span style='white-space:break-spaces'>     </span>B</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 40D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        IReadOnlyList<HtmlRenderText> text = HtmlRenderTestDriver.Render(html, options).Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .OrderBy(visual => visual.PaintOrder)
            .ToList();

        Assert.Equal("A     B", string.Concat(text.Select(visual => visual.Text)));
        Assert.True(text.Select(visual => Math.Round(visual.Y, 3)).Distinct().Count() >= 2);
        Assert.Contains(text, visual => visual.Text.Length > 0 && (char.IsWhiteSpace(visual.Text[0]) || char.IsWhiteSpace(visual.Text[visual.Text.Length - 1])));
    }

    [Fact]
    public void HtmlText_DescendantNoWrapMovesAsOneUnbreakableRange() {
        const string html = "<p style='margin:0;width:100px;font:16px Arial;line-height:18px'>prefix <span style='white-space:nowrap'>first second</span></p>";

        IReadOnlyList<HtmlRenderText> text = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 120D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        }).Pages[0].Visuals.OfType<HtmlRenderText>().ToList();

        HtmlRenderText noWrap = Assert.Single(text, visual => visual.Text == "first second");
        Assert.True(noWrap.Y > text.First(visual => visual.Text.Contains("prefix", StringComparison.Ordinal)).Y);
    }

    [Fact]
    public void HtmlText_BreakSpacesRetainsPreservedRunsBesideFloats() {
        const string html = "<p style='margin:0;width:36px;font:16px Arial;line-height:18px'><span style='float:left;width:12px;height:12px'></span>A<span style='white-space:break-spaces'>     </span>B</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 50D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        IReadOnlyList<HtmlRenderText> text = HtmlRenderTestDriver.Render(html, options).Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .OrderBy(visual => visual.PaintOrder)
            .ToList();

        Assert.Equal("A     B", string.Concat(text.Select(visual => visual.Text)));
        Assert.True(text.Select(visual => Math.Round(visual.Y, 3)).Distinct().Count() >= 2);
    }
}
