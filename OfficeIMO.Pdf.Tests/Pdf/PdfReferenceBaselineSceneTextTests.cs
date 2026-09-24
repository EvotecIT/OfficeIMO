using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReferenceBaselineSceneTextTests {
    // Microsoft Word writes word-final spaces inside TJ elements such as "(e )". The visual scene
    // must position only the painted glyphs, so its reconstructed words match logical extraction.
    [Theory]
    [InlineData("microsoft-word-16.109-native-word-report.pdf")]
    [InlineData("microsoft-word-windows-word-business-delivery-summary.pdf")]
    public void MicrosoftWordReferenceScenePositionsEveryWordAtItsPaintedAdvance(string fileName) {
        string path = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Pdf", "ReferenceBaselines", fileName);
        var document = PdfDocument.Load(path);
        OfficeDrawing drawing = document.Render.Drawing(1);
        Assert.NotEmpty(drawing.Fonts.Faces);

        var runs = new List<(double X, double Baseline, double Advance, double Size, string Text)>();
        CollectTextRuns(drawing, 0D, 0D, runs);
        Assert.NotEmpty(runs);

        var extractedWords = new HashSet<string>(
            PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0].ExtractText()
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries),
            StringComparer.Ordinal);
        var sceneWords = new List<string>();
        foreach (var line in runs.GroupBy(run => Math.Round(run.Baseline))) {
            var ordered = line.OrderBy(run => run.X).ToList();
            var text = new System.Text.StringBuilder(ordered[0].Text);
            for (int index = 1; index < ordered.Count; index++) {
                var previous = ordered[index - 1];
                double gap = ordered[index].X - (previous.X + previous.Advance);
                if (gap > previous.Size * 0.12D) text.Append(' ');
                text.Append(ordered[index].Text);
            }
            sceneWords.AddRange(text.ToString().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        }

        Assert.True(sceneWords.Count > 40, $"Expected a populated page scene, found {sceneWords.Count} words.");
        Assert.Empty(sceneWords.Where(word => !extractedWords.Contains(word)).Distinct());
    }

    // These embedded TrueType programs cannot resolve Unicode scene text by themselves: simple fonts
    // keyed by PDF character code through a (1,0) or (3,0) cmap, a subset with an empty OS/2 table
    // at end of file, and CIDFontType2 subsets that select glyphs through CIDToGIDMap with a symbolic
    // cmap or no cmap at all. Each run must still resolve through its own embedded face.
    [Theory]
    [InlineData("Pdf/ReferenceBaselines/microsoft-word-16.109-native-word-report.pdf")]
    [InlineData("Pdf/ReferenceBaselines/microsoft-word-windows-word-business-delivery-summary.pdf")]
    [InlineData("Pdf/Fixtures/Interoperability/verapdf-tounicode-pass-a.pdf")]
    [InlineData("Pdf/Fixtures/Interoperability/verapdf-tounicode-pass-j.pdf")]
    [InlineData("Pdf/Fixtures/Interoperability/verapdf-optional-content.pdf")]
    [InlineData("Pdf/Fixtures/Fonts/symbolic-truetype-cmap.pdf")]
    [InlineData("../OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf")]
    [InlineData("../OfficeIMO.TestAssets/MultilingualLayout/rtl-0-native.pdf")]
    [InlineData("Pdf/Fixtures/ShapedText/chrome-arabic.pdf")]
    public void CodeKeyedEmbeddedTrueTypeFontsResolveSceneTextThroughTheEmbeddedFace(string relativePath) {
        string path = Path.GetFullPath(Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), relativePath));
        OfficeDrawing drawing = PdfDocument.Load(path).Render.Drawing(1);
        var texts = new List<OfficeDrawingText>();
        CollectTexts(drawing, texts);

        // Every text run in these files is painted by an embedded subset font, so each run must use that
        // font's drawing family and resolve through the registered face rather than a substitute.
        var runs = texts.Where(text => !string.IsNullOrWhiteSpace(text.Text)).ToList();
        Assert.NotEmpty(runs);
        Assert.Empty(runs
            .Where(text => !EmbeddedSubsetFamily.IsMatch(text.Font.FamilyName ?? string.Empty) ||
                !drawing.Fonts.TryResolveFaceForText(text.Text.Trim(), text.Font.FamilyName, text.Font.Style, out OfficeFontFace? face) ||
                !string.Equals(face!.FamilyName, text.Font.FamilyName, StringComparison.Ordinal))
            .Select(text => text.Font.FamilyName + ": " + text.Text)
            .Distinct());
    }

    private static readonly System.Text.RegularExpressions.Regex EmbeddedSubsetFamily =
        new(@"^[A-Z]{6}\+.+-[0-9a-f]{24}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static void CollectTexts(OfficeDrawing drawing, List<OfficeDrawingText> texts) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText text) texts.Add(text);
            else if (element is OfficeDrawingGroup group) CollectTexts(group.Drawing, texts);
            else if (element is OfficeDrawingEffectGroup effectGroup) CollectTexts(effectGroup.Drawing, texts);
        }
    }

    private static void CollectTextRuns(OfficeDrawing drawing, double offsetX, double offsetY,
        List<(double X, double Baseline, double Advance, double Size, string Text)> runs) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingText text:
                    Assert.False(text.HasFrameTransform, $"Unexpected transformed run '{text.Text}'.");
                    // A run without its PDF advance falls back to an estimated, ellipsizing label box.
                    Assert.True(text.TextAdvanceWidth.HasValue, $"Run '{text.Text}' lost its painted PDF advance.");
                    if (string.IsNullOrWhiteSpace(text.Text)) break;
                    runs.Add((offsetX + text.X, offsetY + text.Y + text.Font.Size, text.TextAdvanceWidth!.Value, text.Font.Size, text.Text));
                    break;
                case OfficeDrawingGroup group:
                    Assert.Null(group.FrameTransform);
                    CollectTextRuns(group.Drawing, offsetX + group.X + group.ContentOffsetX, offsetY + group.Y + group.ContentOffsetY, runs);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    Assert.Equal(1D, effectGroup.Transform.M11);
                    Assert.Equal(1D, effectGroup.Transform.M22);
                    CollectTextRuns(effectGroup.Drawing, offsetX + effectGroup.Transform.OffsetX, offsetY + effectGroup.Transform.OffsetY, runs);
                    break;
            }
        }
    }
}
