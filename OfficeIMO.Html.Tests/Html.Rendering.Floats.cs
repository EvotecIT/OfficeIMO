using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlFloats_ApplyLayoutDepthBeforeScanningNestedDescendants() {
        string html = "<span>" + string.Concat(Enumerable.Repeat("<span>", 12))
            + "<span style='float:left'>Float</span>"
            + string.Concat(Enumerable.Repeat("</span>", 12)) + "</span>";

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxLayoutDepth = 8 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.DepthLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutDepth), exception.LimitSource);
    }

    [Fact]
    public void HtmlFloatLeft_WrapsLineBandsAndRestoresFullWidthBelowFloat() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'>"
            + "<span id='left-float' style='float:left;width:30px;height:20px;background:#ff0000'></span>"
            + "One two three four five six seven eight nine ten eleven twelve thirteen fourteen</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape floating = FindPositionedShape(rendered, "span#left-float");
        IReadOnlyList<HtmlRenderText> lines = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();
        Assert.Equal(0D, floating.X, 3);
        Assert.Equal(0D, floating.Y, 3);
        Assert.All(lines.Where(line => line.Y < floating.Y + floating.Height - 0.001D), line => Assert.True(line.X >= floating.X + floating.Width - 0.001D));
        Assert.Contains(lines, line => line.Y >= floating.Y + floating.Height - 0.001D && line.X < 0.001D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlFloatRight_WrapsTextAgainstRightEdge() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px;direction:rtl'>"
            + "<span id='right-float' style='float:inline-start;width:25px;height:20px;background:#0000ff'></span>"
            + "One two three four five six seven eight nine ten</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape floating = FindPositionedShape(rendered, "span#right-float");
        Assert.Equal(75D, floating.X, 3);
        Assert.All(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Where(line => line.Y < 20D - 0.001D),
            line => Assert.True(line.X + line.Width <= floating.X + 0.001D));
    }

    [Fact]
    public void HtmlFloats_PackOnTheSameSideWhenTheyFit() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'>"
            + "<span id='first-float' style='float:left;width:20px;height:20px;background:#ff0000'></span>"
            + "<span id='second-float' style='float:left;width:25px;height:20px;background:#00ff00'></span>"
            + "One two three four five six seven eight</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape first = FindPositionedShape(rendered, "span#first-float");
        HtmlRenderShape second = FindPositionedShape(rendered, "span#second-float");
        HtmlRenderText firstLine = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().OrderBy(text => text.Y).First();
        Assert.Equal(0D, first.X, 3);
        Assert.Equal(first.X + first.Width, second.X, 3);
        Assert.Equal(first.Y, second.Y, 3);
        Assert.True(firstLine.X >= second.X + second.Width - 0.001D);
    }

    [Fact]
    public void HtmlFloatClear_AdvancesOnlyPastTheRequestedSide() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'>"
            + "<span id='left-tall' style='float:left;width:20px;height:30px;background:#ff0000'></span>"
            + "<span id='right-short' style='float:right;width:20px;height:10px;background:#0000ff'></span>"
            + "<span id='clear-right' style='float:right;clear:right;width:15px;height:10px;background:#00ff00'></span>"
            + "One two three four five six seven eight</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape left = FindPositionedShape(rendered, "span#left-tall");
        HtmlRenderShape right = FindPositionedShape(rendered, "span#right-short");
        HtmlRenderShape cleared = FindPositionedShape(rendered, "span#clear-right");
        Assert.Equal(right.Y + right.Height, cleared.Y, 3);
        Assert.True(cleared.Y < left.Y + left.Height);
        Assert.Equal(100D - cleared.Width, cleared.X, 3);
    }

    [Fact]
    public void HtmlFloat_NestedInsideInlineContentStillParticipatesInBlockFlow() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'><span>"
            + "<span id='nested-float' style='float:left;width:30px;height:20px;background:#ff0000'></span>"
            + "Nested float text wraps around the measured box and continues below it.</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape floating = FindPositionedShape(rendered, "span#nested-float");
        HtmlRenderText firstLine = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().OrderBy(text => text.Y).First();
        Assert.Equal(0D, floating.X, 3);
        Assert.True(firstLine.X >= floating.X + floating.Width - 0.001D);
    }

    [Fact]
    public void HtmlFloat_DescendantNoWrapMovesBelowTheObstructedLineAsOneRange() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'>"
            + "<span id='float' style='float:left;width:45px;height:20px;background:#ff0000'></span>"
            + "Lead <span style='white-space:nowrap'><span>No wrap</span> range</span> tail</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape floating = FindPositionedShape(rendered, "span#float");
        IReadOnlyList<HtmlRenderText> text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();
        HtmlRenderText noWrapStart = Assert.Single(text, visual => visual.Text == "No wrap");
        HtmlRenderText noWrapEnd = Assert.Single(text, visual => visual.Text == " range");

        Assert.True(noWrapStart.Y >= floating.Y + floating.Height - 0.001D);
        Assert.Equal(0D, noWrapStart.X, 3);
        Assert.Equal(noWrapStart.Y, noWrapEnd.Y, 3);
        Assert.True(noWrapEnd.X > noWrapStart.X);
    }

    [Fact]
    public void HtmlFloat_AutomaticHyphenationUsesAuthoredBreakpointsAndCharacter() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D),
            TextHyphenationCallback = token => token == "typography" ? new[] { 2, 5, 7 } : Array.Empty<int>()
        };
        const string html = "<p style='width:100px;margin:0;font-size:12px;line-height:14px'>"
            + "<span style='float:left;width:50px;height:28px'></span>"
            + "<span style='hyphens:auto;hyphenate-character:\"·\"'>typography</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderText> fragments = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();

        Assert.Contains(fragments, fragment => fragment.Text.EndsWith("·", StringComparison.Ordinal));
        Assert.Equal("typography", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))));
    }

    [Fact]
    public void HtmlFloat_ExpandsPreformattedTabsUsingAuthoredTabStops() {
        const string compact = "<pre style='width:120px;margin:0;font:12px/14px Consolas;tab-size:2'><span style='float:left;width:20px;height:14px'></span>A\tB</pre>";
        const string wide = "<pre style='width:120px;margin:0;font:12px/14px Consolas;tab-size:8'><span style='float:left;width:20px;height:14px'></span>A\tB</pre>";

        HtmlRenderText[] compactText = HtmlRenderTestDriver.Render(compact).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] wideText = HtmlRenderTestDriver.Render(wide).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.DoesNotContain('\t', string.Concat(compactText.Select(text => text.Text)));
        Assert.DoesNotContain('\t', string.Concat(wideText.Select(text => text.Text)));
        Assert.True(Assert.Single(wideText, text => text.Text == "B").X > Assert.Single(compactText, text => text.Text == "B").X);
    }

    [Fact]
    public void HtmlFloat_PreservesInteriorTabPositionsInsideSpacedWhitespaceRuns() {
        const string html = "<pre style='width:140px;margin:0;font:12px/14px Consolas;tab-size:8'><span style='float:left;width:20px;height:14px'></span>A \t B</pre>";

        HtmlRenderText[] text = HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText before = Assert.Single(text, item => item.Text == "A ");
        HtmlRenderText after = Assert.Single(text, item => item.Text == " B");

        Assert.True(after.X > before.X + before.Width + 15D);
    }

    [Fact]
    public void HtmlFloat_BreakAllSuppressesManualAndAutomaticHyphenation() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D),
            TextHyphenationCallback = _ => throw new InvalidOperationException("break-all must suppress automatic hyphenation callbacks")
        };
        const string html = "<p style='width:100px;margin:0;font-size:12px;line-height:14px'>"
            + "<span style='float:left;width:50px;height:28px'></span>"
            + "<span style='word-break:break-all;hyphens:auto;hyphenate-character:\"·\"'>ty\u00ADpography</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderText[] fragments = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.True(fragments.Select(fragment => fragment.Y).Distinct().Count() > 1);
        Assert.InRange(fragments[0].X, 49.9D, 50.1D);
        Assert.InRange(fragments[0].Y, -0.001D, 0.001D);
        Assert.DoesNotContain("·", string.Concat(fragments.Select(fragment => fragment.Text)), StringComparison.Ordinal);
        Assert.Equal("typography", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))));
    }

    [Fact]
    public void HtmlFloat_LineClampTruncatesFloatAwareLinesAndAddsEllipsis() {
        const string html = "<p style='width:100px;margin:0;font-size:12px;line-height:14px;overflow:hidden;line-clamp:2'>"
            + "<span style='float:left;width:30px;height:28px'></span>one two three four five six seven eight nine ten eleven twelve</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> lines = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToList();

        Assert.Equal(2, lines.Select(line => line.Y).Distinct().Count());
        Assert.EndsWith("\u2026", lines[lines.Count - 1].Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlFloat_NoWrapTextOverflowUsesEllipsisBesideAnObstruction() {
        const string html = "<p style='width:100px;margin:0;font-size:12px;line-height:14px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis'>"
            + "<span style='float:left;width:30px;height:20px'></span>one two three four five six seven eight</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> fragments = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .OrderBy(fragment => fragment.X)
            .ToList();

        Assert.Single(fragments.Select(fragment => fragment.Y).Distinct());
        Assert.EndsWith("\u2026", fragments[fragments.Count - 1].Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlFloat_CjkUsesPreferredLineBreaksBesideAnObstruction() {
        const string text = "日本語中文";
        string html = """
            <div style="width:70px;font:12px/14px Arial,sans-serif">
              <span style="float:left;width:38px;height:28px"></span>日本語中文
            </div>
            """;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        });

        IReadOnlyList<HtmlRenderText> lines = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .Where(visual => text.Contains(visual.Text, StringComparison.Ordinal))
            .OrderBy(visual => visual.Y)
            .ThenBy(visual => visual.X)
            .ToList();
        Assert.True(lines.Select(line => line.Y).Distinct().Count() > 1);
        Assert.Equal(text, string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))));
    }

    [Fact]
    public void HtmlFloatImage_UsesIntrinsicAspectRatioWithoutConsumingTheLine() {
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";
        string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'>"
            + "<img id='float-image' src='data:image/png;base64," + png + "' style='float:left;height:20px'>"
            + "Image text wraps beside the intrinsic image and then returns to full width.</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), visual => visual.Source == "img#float-image");
        HtmlRenderText firstLine = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().OrderBy(text => text.Y).First();
        Assert.Equal(20D, image.Width, 3);
        Assert.Equal(20D, image.Height, 3);
        Assert.True(firstLine.X >= image.X + image.Width - 0.001D);
    }

    [Fact]
    public void HtmlFloat_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<p style='width:80px;margin:0;font-size:10px;line-height:10px'>"
            + "<span id='float-paint' style='float:left;width:20px;height:20px;background:#ff0000'></span>"
            + "FloatPdfMarker wraps beside the box.</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(80D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(5, 5));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"0\" y=\"0\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("FloatPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlFloat_PaginatesWrappedLinesWithoutRepeatingTheFloat() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:10px'>"
            + "<span id='paged-float' style='float:left;width:30px;height:20px;background:#ff0000'></span>"
            + "One two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen twenty.</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.True(rendered.Pages.Count >= 2);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "span#paged-float");
        Assert.All(
            rendered.Pages.Skip(1),
            page => Assert.DoesNotContain(page.Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "span#paged-float"));
        Assert.Contains(rendered.Pages.Skip(1).SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.Length > 0);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlFloat_InvalidValuesUseCatalogedDiagnostics() {
        const string html = "<p style='margin:0'><span id='invalid-float' style='float:up;clear:around'>Text</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
        Assert.Equal("span#invalid-float", diagnostic.Source);
        Assert.Contains("float=up", diagnostic.Detail);
        Assert.Contains("clear=around", diagnostic.Detail);
        Assert.Contains(HtmlRenderDiagnosticCodes.FloatValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.FloatValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(float:left)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clear:both)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(float:up)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clear:around)"));
    }
}
