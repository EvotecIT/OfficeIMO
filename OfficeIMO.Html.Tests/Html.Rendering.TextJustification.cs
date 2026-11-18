using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string JustifiedText =
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.";

    public static TheoryData<string, string> JustifiedParagraphCases => new() {
        // The issue's own markup.
        { "<p style='text-align: justify; width: 300px; margin:0'>" + JustifiedText + "</p>", "start" },
        // Inherited through block, list item and table cell boxes.
        { "<div style='text-align:justify;width:300px'><p style='margin:0'>" + JustifiedText + "</p></div>", "start" },
        { "<div style='text-align:justify;width:300px'><ul style='margin:0;padding:0;list-style:none'><li>" + JustifiedText + "</li></ul></div>", "start" },
        { "<div style='text-align:justify'><table style='width:300px;border-spacing:0'><tr><td style='padding:0'>" + JustifiedText + "</td></tr></table></div>", "start" },
        // The legacy presentational attribute maps to the same text-align value.
        { "<p align='justify' style='width:300px;margin:0'>" + JustifiedText + "</p>", "start" },
        { "<p style='text-align:justify;text-align-last:center;width:300px;margin:0'>" + JustifiedText + "</p>", "center" },
        { "<p style='text-align:justify;text-align-last:end;width:300px;margin:0'>" + JustifiedText + "</p>", "end" },
        { "<p style='text-align:justify-all;width:300px;margin:0'>" + JustifiedText + "</p>", "justified" },
        { "<p style='text-align:left;text-align-last:justify;width:300px;margin:0'>" + JustifiedText + "</p>", "justified-last-only" },
        // A text-align-last after justify-all wins the cascade, as with the shorthand expanded.
        { "<p style='text-align:justify-all;text-align-last:auto;width:300px;margin:0'>" + JustifiedText + "</p>", "start" },
        // match-parent takes the parent's justify value, then applies its own text-justify.
        { "<div style='text-align:justify;text-justify:none;width:300px'><p style='margin:0;text-align:match-parent;text-justify:inter-word'>" + JustifiedText + "</p></div>", "start" },
        // Inheriting justify-all also inherits the parent's computed last-line alignment.
        { "<div style='text-align:justify-all;text-align-last:center;width:300px'><p style='margin:0'>" + JustifiedText + "</p></div>", "center" },
        { "<div style='text-align:justify-all;width:300px'><p style='margin:0;text-align:match-parent'>" + JustifiedText + "</p></div>", "justified" },
        // An explicit reset is not inherited.
        { "<div style='text-align:justify;text-align-last:center;width:300px'><p style='margin:0;text-align-last:initial'>" + JustifiedText + "</p></div>", "start" }
    };

    [Theory]
    [MemberData(nameof(JustifiedParagraphCases))]
    public void HtmlRenderer_JustifiesWrappedLinesAndAlignsTheLastLine(string html, string lastLine) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<JustifiedLine> lines = JustifiedLines(rendered);

        Assert.True(lines.Count >= 3, "The fixture must wrap onto at least three lines.");
        JustifiedLine[] earlier = lines.Take(lines.Count - 1).ToArray();
        Assert.All(earlier, line => Assert.Equal(0D, line.Left, 1));
        if (lastLine == "justified-last-only") {
            Assert.Contains(earlier, line => line.Right < 299D);
        } else {
            Assert.All(earlier, line => Assert.Equal(300D, line.Right, 1));
        }
        JustifiedLine last = lines[lines.Count - 1];
        switch (lastLine) {
            case "start":
                Assert.Equal(0D, last.Left, 1);
                Assert.True(last.Right < 280D, "The last line of justified text must stay at start.");
                break;
            case "center":
                Assert.Equal(300D - last.Right, last.Left, 1);
                Assert.True(last.Left > 10D);
                break;
            case "end":
                Assert.Equal(300D, last.Right, 1);
                Assert.True(last.Left > 10D);
                break;
            default:
                Assert.Equal(0D, last.Left, 1);
                Assert.Equal(300D, last.Right, 1);
                break;
        }
    }

    [Fact]
    public void HtmlRenderer_JustifiedLinesEndingInForcedBreaksStayAtStart() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:justify;width:300px;margin:0'>Forced break here<br>" + JustifiedText + "</p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });
        IReadOnlyList<JustifiedLine> lines = JustifiedLines(rendered);

        Assert.Equal("Forced break here", lines[0].Text);
        Assert.Equal(0D, lines[0].Left, 1);
        Assert.True(lines[0].Right < 200D);
        Assert.Equal(300D, lines[1].Right, 1);
    }

    [Fact]
    public void HtmlRenderer_JustifiesRightToLeftLinesAndKeepsTheLastLineAtRightStart() {
        const string hebrew = "שלום עולם זהו טקסט בעברית שאמור להיות מיושר לשני הצדדים בפסקה ארוכה מספיק כדי לגלוש לשורה שנייה ושלישית ורביעית.";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p dir='rtl' style='text-align:justify;width:300px;margin:0'>" + hebrew + "</p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });
        IReadOnlyList<JustifiedLine> lines = JustifiedLines(rendered);

        Assert.True(lines.Count >= 2);
        foreach (JustifiedLine line in lines.Take(lines.Count - 1)) {
            Assert.Equal(0D, line.Left, 1);
            Assert.Equal(300D, line.Right, 1);
        }
        Assert.Equal(300D, lines[lines.Count - 1].Right, 1);
        Assert.True(lines[lines.Count - 1].Left > 10D);
    }

    [Fact]
    public void HtmlRenderer_TextJustifyNoneKeepsJustifiedTextAtStart() {
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) };
        IReadOnlyList<JustifiedLine> none = JustifiedLines(HtmlRenderTestDriver.Render(
            "<div style='text-align:justify;text-justify:none'><p style='width:300px;margin:0'>" + JustifiedText + "</p></div>", options));
        IReadOnlyList<JustifiedLine> start = JustifiedLines(HtmlRenderTestDriver.Render(
            "<div style='text-align:start'><p style='width:300px;margin:0'>" + JustifiedText + "</p></div>", options));

        Assert.Equal(start.Select(line => (line.Text, Math.Round(line.Left, 2), Math.Round(line.Right, 2))),
            none.Select(line => (line.Text, Math.Round(line.Left, 2), Math.Round(line.Right, 2))));
    }

    [Fact]
    public void HtmlRenderer_TextJustifyNoneKeepsTheSpacesOfItsOwnText() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align-last:justify;width:300px;margin:0'>alpha <span style='text-justify:none'>one two</span> omega</p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });
        JustifiedLine line = Assert.Single(JustifiedLines(rendered));

        Assert.Equal(300D, line.Right, 1);
        Assert.Contains(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(), text => text.Text == "one two");
    }

    [Fact]
    public void HtmlRenderer_RightToLeftLinesWithInlineBoxesKeepLeftToRightWordOrder() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p dir='rtl' style='text-align:justify;width:300px;margin:0'>שלום <span style='display:inline-block;width:10px;height:10px'></span> "
            + "one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen</p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });

        // Without bidi resolution the line keeps the phrase whole, so its words stay in order.
        Assert.Contains(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(), text => text.Text.Contains("one two"));
    }

    [Fact]
    public void HtmlRenderer_ReportsInterCharacterJustificationAsWordSpacing() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:justify;text-justify:inter-character;width:300px;margin:0'>" + JustifiedText + "</p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });

        Assert.Equal(300D, JustifiedLines(rendered)[0].Right, 1);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.NativeDiagnostics,
            item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported && item.Detail == "text-justify=inter-character");
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
    }

    [Fact]
    public void HtmlRenderer_JustifiedInlineBackgroundCoversTheAddedWordSpace() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:justify;width:300px;margin:0'><span style='background:#ffe08a'>" + JustifiedText + "</span></p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });
        int lineCount = JustifiedLines(rendered).Count;
        HtmlRenderShape[] backgrounds = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderShape>()
            .Where(shape => shape.Shape.FillColor is OfficeColor fill && fill.R == 0xFF && fill.G == 0xE0)
            .ToArray();

        // One fragment per line, not one per word: the added space belongs to the inline box.
        Assert.Equal(lineCount, backgrounds.Length);
        HtmlRenderShape first = backgrounds.OrderBy(shape => shape.Y).First();
        Assert.Equal(0D, first.X, 1);
        Assert.Equal(300D, first.X + first.Width, 1);
    }

    [Fact]
    public void HtmlToPdf_JustifiedParagraphKeepsSingleSpacesInExtractedText() {
        byte[] pdf = HtmlConversionDocument.Parse("<p style='text-align: justify; width: 300px;'>" + JustifiedText + "</p>")
            .ToPdf(new HtmlPdfExportOptions {
                AllowSystemFontFallback = false,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
            }).ToBytes();
        string extracted = PdfCore.PdfDocument.Load(new System.IO.MemoryStream(pdf)).ExtractText();

        Assert.Equal(JustifiedText.Split(' '), extracted.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        Assert.DoesNotContain("  ", extracted.Replace("\r", string.Empty).Replace("\n", " "), StringComparison.Ordinal);
    }

    private sealed record JustifiedLine(double Left, double Right, string Text);

    // Word extents per line box, ignoring separators so start, end and justified lines can be compared.
    private static IReadOnlyList<JustifiedLine> JustifiedLines(HtmlRenderDocument rendered) =>
        JustificationVisuals(rendered.Pages[0].Visuals)
            .OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text))
            .GroupBy(text => Math.Round(text.LayoutY, 1))
            .OrderBy(group => group.Key)
            .Select(group => new JustifiedLine(
                group.Min(text => text.X),
                group.Max(text => text.X + (text.TextAdvanceWidth ?? text.Width)),
                string.Join(" ", group.OrderBy(text => text.X).Select(text => text.Text.Trim()))))
            .ToArray();

    private static IEnumerable<HtmlRenderVisual> JustificationVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in JustificationVisuals(children)) yield return child;
        }
    }
}
