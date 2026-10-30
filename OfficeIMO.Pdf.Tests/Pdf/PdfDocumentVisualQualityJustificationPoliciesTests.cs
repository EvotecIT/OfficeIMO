using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ParagraphJustificationKeepsExplicitBreakAndFinalLinePoliciesIndependent(bool explicitBreak, bool finalLine) {
        var style = new PdfParagraphStyle { JustifyExplicitLineBreaks = explicitBreak, JustifyLastLine = finalLine };
        byte[] bytes = PdfDocument.Create(JustificationPolicyOptions()).Paragraph(paragraph => paragraph
            .Text("Alpha beta gamma").LineBreak().Text("Short final words"), PdfAlign.Justify, style: style).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        List<List<Letter>> lines = GetNonWhitespaceLetterLines(pdf.GetPage(1));
        Assert.Equal(2, lines.Count);
        AssertJustificationPolicyWidth(lines[0], explicitBreak);
        AssertJustificationPolicyWidth(lines[1], finalLine);
    }

    [Fact]
    public void ParagraphJustificationCapturesCharacterAndLinePoliciesBeforeCallerMutation() {
        var style = CharacterJustificationPolicy();
        var document = PdfDocument.Create(JustificationPolicyOptions()).Paragraph(paragraph => paragraph
            .Text("Alpha beta gamma").LineBreak().Text("Short final words"), PdfAlign.Justify, style: style);
        style.JustificationMode = PdfJustificationMode.Words;
        style.JustifyExplicitLineBreaks = false;
        style.JustifyLastLine = false;
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        foreach (List<Letter> line in GetNonWhitespaceLetterLines(pdf.GetPage(1))) {
            AssertJustificationPolicyWidth(line, expanded: true);
            Assert.True(line[1].StartBaseLine.X - line[0].EndBaseLine.X > 3D);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CharacterJustificationKeepsCombiningClustersTogetherAcrossFormattingRuns(bool splitRuns) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B', 0x0301);
        var options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Cluster Font",
            font));
        PdfTextRun[] runs = splitRuns ? new[] { new PdfTextRun("A"), new PdfTextRun("\u0301B", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(1, 0, 0)) }
            : new[] { new PdfTextRun("A\u0301B") };
        byte[] controlBytes = RenderJustificationPolicyRuns(options, runs, distributed: false);
        byte[] actualBytes = RenderJustificationPolicyRuns(options, runs, distributed: true);
        Letter[] control = ReadJustificationPolicyLetters(controlBytes);
        Letter[] actual = ReadJustificationPolicyLetters(actualBytes);
        Assert.Equal("A\u0301B", string.Concat(actual.Select(letter => letter.Value)));
        Assert.Equal(string.Concat(control.Select(letter => letter.Value)), string.Concat(actual.Select(letter => letter.Value)));
        if (splitRuns) {
            Assert.Equal(3, actual.Length);
            Assert.Equal(control[1].StartBaseLine.X - control[0].StartBaseLine.X,
                actual[1].StartBaseLine.X - actual[0].StartBaseLine.X, 3);
        } else {
            var program = PdfTrueTypeFontProgram.Parse(font, "Cluster Font");
            Assert.True(program.TryGetGlyphId('A', out int baseGlyph));
            Assert.True(program.TryGetGlyphId(0x0301, out int markGlyph));
            Assert.True(program.TryGetGlyphId('B', out int lastGlyph));
            string cluster = baseGlyph.ToString("X4") + markGlyph.ToString("X4");
            string[] controlShows = JustificationShownGlyphHex(controlBytes);
            string[] actualShows = JustificationShownGlyphHex(actualBytes);
            Assert.Equal(cluster + lastGlyph.ToString("X4"), string.Concat(controlShows));
            Assert.Equal(string.Concat(controlShows), string.Concat(actualShows));
            // One uninterrupted encoded cluster has the same font advances in both artifacts.
            // PdfPig can project that ActualText cluster into a single logical Letter.
            Assert.Contains(actualShows, hex => hex.Contains(cluster, StringComparison.Ordinal));
        }
        Assert.True(actual.Last().StartBaseLine.X - actual.First().StartBaseLine.X >
            control.Last().StartBaseLine.X - control.First().StartBaseLine.X + 100D);
        AssertJustificationPolicyWidth(actual.ToList(), expanded: true);
    }

    [Fact]
    public void CharacterJustificationKeepsShapedLigaturesAsSingleDistributionUnits() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn");
        PdfCreateOptions options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Ligature Font",
            font));
        options.TextShapingMode = PdfTextShapingMode.OpenTypeLigatures;
        PdfGlyphRun shaped = PdfTrueTypeFontProgram.Parse(font, "Ligature Font").ShapeText("fifi",
            PdfTextShapingOptions.ForRendering("Ligature Font", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(new[] { "fi", "fi" }, shaped.Glyphs.Select(glyph => glyph.UnicodeText));
        var runs = new[] { new PdfTextRun("fifi") };
        byte[] controlBytes = RenderJustificationPolicyRuns(options, runs, distributed: false);
        byte[] actualBytes = RenderJustificationPolicyRuns(options, runs, distributed: true);
        Letter[] control = ReadJustificationPolicyLetters(controlBytes);
        Letter[] actual = ReadJustificationPolicyLetters(actualBytes);
        string expectedGlyphs = PdfGlyphRun.ToGlyphHex(shaped.Glyphs);
        Assert.Equal(expectedGlyphs, string.Concat(JustificationShownGlyphHex(controlBytes)));
        Assert.Equal(expectedGlyphs, string.Concat(JustificationShownGlyphHex(actualBytes)));
        Assert.Equal("fifi", string.Concat(actual.Select(letter => letter.Value)));
        Assert.Equal(control[0].Width, actual[0].Width, 3);
        // Ambiguous ActualText deliberately has no public glyph projection. Inspect the
        // real PDF show: two unchanged CIDs and one adjustment at their sole boundary.
        double adjustment = Assert.Single(JustificationTextAdjustments(actualBytes));
        double natural = shaped.TotalAdvanceWidth1000 * options.DefaultFontSize.ToPoints() / 1000D;
        double spare = options.PageWidth.ToPoints() - options.MarginLeft.ToPoints() - options.MarginRight.ToPoints() - natural;
        Assert.Equal(spare, -adjustment * options.DefaultFontSize.ToPoints() / 1000D, 3);
        AssertJustificationPolicyWidth(actual.ToList(), expanded: true);
    }

    [Fact]
    public void CharacterJustificationAddsPagePointTrackingWithoutChangingAuthoredScalingAndSpacing() {
        PdfCreateOptions options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Spacing Font",
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B', 'C', 'D')));
        var runs = new[] { new PdfTextRun("AB").WithHorizontalTextScaling(150).WithCharacterSpacing(2),
            new PdfTextRun("CD").WithHorizontalTextScaling(75) };
        Letter[] control = ReadJustificationPolicyLetters(RenderJustificationPolicyRuns(options, runs, distributed: false));
        Letter[] actual = ReadJustificationPolicyLetters(RenderJustificationPolicyRuns(options, runs, distributed: true));
        double natural = control.Last().EndBaseLine.X - control.First().StartBaseLine.X;
        double extra = (240D - natural) / 3D;
        Assert.Equal(control.Length, actual.Length);
        for (int index = 0; index < actual.Length; index++) Assert.Equal(control[index].Width, actual[index].Width, 3);
        for (int index = 1; index < actual.Length; index++)
            Assert.InRange(Math.Abs(actual[index].StartBaseLine.X - actual[index - 1].StartBaseLine.X -
                (control[index].StartBaseLine.X - control[index - 1].StartBaseLine.X) - extra), 0D, .02D);
        AssertJustificationPolicyWidth(actual.ToList(), expanded: true);
    }

    [Fact]
    public void CharacterJustificationKeepsAnExplicitTabAnchorAndExpandsOnlyItsTrailingWordSeparators() {
        var style = CharacterJustificationPolicy();
        style.TabStops.Add(new PdfTabStop(120));
        byte[] bytes = PdfDocument.Create(JustificationPolicyOptions()).Paragraph(paragraph => paragraph.Text("Alpha one\tBeta last words"),
            PdfAlign.Justify, style: style).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Letter[] letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Letter beta = letters.Single(letter => letter.Value == "B");
        Assert.Equal(150D, beta.StartBaseLine.X, 3);
        Assert.InRange(letters[1].StartBaseLine.X - letters[0].EndBaseLine.X, -1D, 1D);
        int betaIndex = Array.IndexOf(letters, beta);
        Assert.InRange(letters[betaIndex + 1].StartBaseLine.X - beta.EndBaseLine.X, -1D, 1D);
        Assert.InRange(letters[5].StartBaseLine.X - letters[4].EndBaseLine.X, 2D, 5D);
        Assert.Equal(270D, letters.Last().EndBaseLine.X, 2);
    }

    [Fact]
    public void CellParagraphJustificationRetainsSourceFinalLineIdentityAcrossFlattening() {
        var words = new[] { new PdfTextRun("Alpha beta gamma\nShort final words") };
        var characters = new[] { new PdfTextRun("Second final") };
        var paragraphs = new[] {
            new PdfTableCellParagraph(words, align: PdfAlign.Justify, justifyExplicitLineBreaks: true),
            new PdfTableCellParagraph(characters, align: PdfAlign.Justify,
                justificationMode: PdfJustificationMode.Characters, justifyLastLine: true)
        };
        var cell = new PdfTableCell(words.Concat(characters), paragraphs);
        byte[] bytes = PdfDocument.Create(JustificationPolicyOptions()).Table(new[] { new[] { cell } }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        List<List<Letter>> lines = GetNonWhitespaceLetterLines(pdf.GetPage(1));
        Assert.Equal(3, lines.Count);
        Assert.Equal(lines[2].Last().EndBaseLine.X, lines[0].Last().EndBaseLine.X, 2);
        Assert.True(lines[2].Last().EndBaseLine.X - lines[1].Last().EndBaseLine.X > 50D);
        Assert.InRange(lines[1].Last().EndBaseLine.X - lines[1].First().StartBaseLine.X, 70D, 140D);
        Assert.True(lines[2][1].StartBaseLine.X - lines[2][0].EndBaseLine.X > 3D);
    }

    [Fact]
    public void CharacterJustificationRetainsInlineBoxGeometryWhileDistributingItsNeighbors() {
        byte[] bytes = PdfDocument.Create(JustificationPolicyOptions()).Paragraph(paragraph => paragraph.Runs(new[] {
            new PdfTextRun("A"), PdfTextRun.Inline(new PdfInlineBox(20, 10, background: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0, 1, 0))), new PdfTextRun("B")
        }), PdfAlign.Justify, style: CharacterJustificationPolicy()).ToBytes();
        Letter[] letters = ReadJustificationPolicyLetters(bytes);
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        AssertJustificationPolicyWidth(letters.ToList(), expanded: true);
        var rectangle = Assert.Single(ExtractPaintedRectangles(PdfTestContentText.ReadDecoded(bytes), "0 1 0 rg", "f"));
        Assert.Equal(20D, rectangle.W, 3);
        Assert.Equal(10D, rectangle.H, 3);
        Assert.True(rectangle.X > letters[0].EndBaseLine.X);
        Assert.True(rectangle.X + rectangle.W < letters[1].StartBaseLine.X);
    }

    private static PdfCreateOptions JustificationPolicyOptions() => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(220), MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30),
        MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12), CompressContentStreams = false
    };
    private static PdfParagraphStyle CharacterJustificationPolicy() => new() {
        JustificationMode = PdfJustificationMode.Characters, JustifyExplicitLineBreaks = true, JustifyLastLine = true
    };
    private static byte[] RenderJustificationPolicyRuns(PdfCreateOptions options, PdfTextRun[] runs, bool distributed) =>
        PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(runs), distributed ? PdfAlign.Justify : PdfAlign.Left,
            style: distributed ? CharacterJustificationPolicy() : new PdfParagraphStyle()).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    private static Letter[] ReadJustificationPolicyLetters(byte[] bytes) {
        using var pdf = PdfPigDocument.Open(bytes);
        return pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
    }
    private static void AssertJustificationPolicyWidth(List<Letter> line, bool expanded) {
        double width = line.Last().EndBaseLine.X - line.First().StartBaseLine.X;
        Assert.InRange(width, expanded ? 239.95D : 70D, expanded ? 240.05D : 140D);
    }

    private static string[] JustificationShownGlyphHex(byte[] bytes) {
        return System.Text.RegularExpressions.Regex.Matches(JustificationTextObjects(bytes),
            @"<(?<single>[0-9A-F]+)>\s*Tj\b|\[(?<array>[^\]]*)\]\s*TJ\b")
            .Cast<System.Text.RegularExpressions.Match>().SelectMany(match => match.Groups["single"].Success
                ? new[] { match.Groups["single"].Value }
                : System.Text.RegularExpressions.Regex.Matches(match.Groups["array"].Value, @"<(?<hex>[0-9A-F]+)>")
                    .Cast<System.Text.RegularExpressions.Match>().Select(part => part.Groups["hex"].Value)).ToArray();
    }

    private static double[] JustificationTextAdjustments(byte[] bytes) => System.Text.RegularExpressions.Regex
        .Matches(JustificationTextObjects(bytes), @"\[(?<parts>[^\]]*)\]\s*TJ")
        .Cast<System.Text.RegularExpressions.Match>()
        .SelectMany(match => System.Text.RegularExpressions.Regex.Matches(
            System.Text.RegularExpressions.Regex.Replace(match.Groups["parts"].Value, @"<[^>]*>", " "), @"[-+]?\d*\.?\d+")
            .Cast<System.Text.RegularExpressions.Match>())
        .Select(match => double.Parse(match.Value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();

    private static string JustificationTextObjects(byte[] bytes) => string.Join("\n", System.Text.RegularExpressions.Regex
        .Matches(PdfTestContentText.ReadDecoded(bytes), @"(?s)\bBT\b.*?\bET\b")
        .Cast<System.Text.RegularExpressions.Match>().Select(match => match.Value));
}
