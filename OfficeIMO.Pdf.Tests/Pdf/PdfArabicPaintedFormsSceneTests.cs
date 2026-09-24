using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfArabicPaintedFormsSceneTests {
    [Fact]
    public void CoincidentArabicPaintPassesShareJoiningContext() {
        var spans = new List<PdfTextSpan> {
            new("\u0628", "F1", 12D, 100.6D, 700D, 6D),
            new("\u0628", "F1", 12D, 100D, 700D, 6D),
            new("\u0628", "F1", 12D, 94D, 700D, 6D),
            new("\u0628", "F1", 12D, 94.6D, 700D, 6D)
        };

        PdfArabicPaintedForms.Apply(spans);

        Assert.Equal(spans[0].Text, spans[1].Text);
        Assert.Equal(spans[2].Text, spans[3].Text);
        Assert.Equal("\uFE91", spans[0].Text);
        Assert.Equal("\uFE90", spans[2].Text);
    }

    // cairo maps every contextual glyph back to its base letter. The page scene must name each
    // painted letter by the presentation form its joining context selects, including Persian and
    // Urdu letters and a word whose middle letters are painted in a second font. Expected forms
    // follow the Unicode joining rules; each line is listed in logical (right-to-left) order.
    [Fact]
    public void CairoExtendedArabicSceneNamesContextualPresentationForms() {
        string path = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Pdf", "Fixtures", "ShapedText", "cairo-arabic-extended.pdf");
        OfficeDrawing drawing = PdfDocument.Load(path).Render.Drawing(1);
        var runs = new List<(double X, double Baseline, char Text)>();
        CollectArabicRuns(drawing, 0D, 0D, runs);

        string[] lines = runs
            .GroupBy(run => Math.Round(run.Baseline))
            .OrderBy(line => line.Key)
            .Select(line => new string(line.OrderByDescending(run => run.X).Select(run => run.Text).ToArray()))
            .ToArray();

        Assert.Equal(new[] {
            "\uFB58\uFB8B\uFEED\uFEEB\uFEB6\uFB94\uFEB0\uFE8D\uFEAD\uFEB5\uFB90\uFBFF\uFED4\uFBFF\uFE96\uFB7C\uFE8E\uFB56",
            "\uFBFE\uFBA7\uFB68\uFBFF\uFEB4\uFB67\uFB88\uFBFE\uFB69\uFE8E\uFBA8\uFBAF\uFEE3\uFBFF\uFB9F\uFE91\uFB8D\uFBFC",
            "\uFE8D\uFEDF\uFEE4\uFEAE\uFE8D\uFE9F\uFECC\uFE94\uFE97\uFE92\uFEAA\uFE83"
        }, lines);
    }

    private static void CollectArabicRuns(OfficeDrawing drawing, double offsetX, double offsetY, List<(double X, double Baseline, char Text)> runs) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingText text when text.Text.Length == 1 && IsArabic(text.Text[0]):
                    runs.Add((offsetX + text.X, offsetY + text.Y + text.Font.Size, text.Text[0]));
                    break;
                case OfficeDrawingGroup group:
                    CollectArabicRuns(group.Drawing, offsetX + group.X + group.ContentOffsetX, offsetY + group.Y + group.ContentOffsetY, runs);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    CollectArabicRuns(effectGroup.Drawing, offsetX + effectGroup.Transform.OffsetX, offsetY + effectGroup.Transform.OffsetY, runs);
                    break;
            }
        }
    }

    private static bool IsArabic(char value) =>
        value >= '\u0600' && value <= '\u06FF' || value >= '\uFB50' && value <= '\uFDFF' || value >= '\uFE70' && value <= '\uFEFF';
}
