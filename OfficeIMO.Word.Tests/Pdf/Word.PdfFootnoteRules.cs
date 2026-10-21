using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    private const string NativeCalibriRegularHash = "57e93ff143572bcc092c1f56202d43ddc07bb119aa104c6d7af163470fb551f6";
    private static readonly IReadOnlyDictionary<string, (string Input, string Reference)> FootnoteRuleFixtureHashes =
        new Dictionary<string, (string Input, string Reference)>(StringComparer.Ordinal) {
            ["geometry-indents-alignment-mixed"] = ("579421afd890a3c5f161d10bb39d91ce6b71914cd564a1fbc877d02fa964363b", "736fdd4795c739edd782726c81921834f913673fa65675cb7914095c3b11884f"),
            ["paint-font-size-style"] = ("648af103041c75a26764205f8d36cc30adb8689dc96214ef2f2be3f52b83c86d", "fc114d792f31f61e9dd99417893afc8e0b2d072781449be855795a03adf1e8ae"),
            ["continuation-notice-compat14"] = ("790f0fdf7e332789b05595c5bb2670c0d000686046e37c21ca42ea660a49f424", "d2ad542b4e61ff3e204eab9bfc277c32028e7c7352ac694c882b337f65b88bea"),
            ["continuation-notice-compat15"] = ("7fcc2f8ecd3d1327807e0eeb7e230155039b1cde9cac15e084a2ac051856adfe", "8201ac28be10750aafdffa64c0a782ab9367da3176569346c6dd6e6d2aadde36"),
            ["continuation-plus-new-note"] = ("bb77e1f4a4e95d4a6c820ca21f4699519140149e8a79d5a748386c7fbcdeb7bf", "cc91726cf797fec5acf69f0573d5ef1ce2b1bdfd510dab7bda34f769c414b662"),
            ["continuation-notice-compat14-font6"] = ("b01b8be290637863a972863cdfaeecc0ce304e163c4f5a7f9b587242eabdef14", "2c7cc823631aa693c0556feb77cc188376e3963d05e6b85789884ce4e606f498"),
            ["continuation-notice-compat14-font18"] = ("24b361dafbb87203c8c8d836607ef1c4b0044fc5833d6f657217eb6586fdd398", "3c0221df6ab3b6ddf2a68512cb6ae39a0d2e400cd72eab7111dc655bf857a94e"),
            ["continuation-notice-compat14-short-note"] = ("c23fbdd2c21dd97af0f44d3a25dd9ee56cced16c2b9eb03aa8f8024b349cf462", "c2cb0c93957e548e9730b4e51b82448bdad7bbb506a62f7791eb3988282e854d"),
            ["continuation-notice-compat15-short-note"] = ("2708a49684307868eabdfe099ade7fb43e639c60b9b29bf7a207fcfb16b83a4d", "7fb80143c63c67ea3e1ca3b8859d43329fd6f61b74c54ed02dd92580bf8ab9c2")
        };

    [Fact]
    public void SaveAsPdf_FootnoteRulesRespectIndentedFramesAndMixedTextOrder() {
        PageFootnoteArtifact bytes = ConvertFootnoteRuleFixture("geometry-indents-alignment-mixed");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        Page page = actual.GetPage(1);
        Letter[] lead = PageFootnoteToken(page, "LEAD");
        double noteBaseline = PageFootnoteToken(page, "SHORT NOTE BODY")[0].StartBaseLine.Y;
        // REFERENCE TAIL also occurs in the body. Select the authored mixed
        // paragraph's following line between LEAD and its note, excluding the
        // smaller note label, instead of matching that unrelated body token.
        Letter[] tail = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value) &&
            letter.StartBaseLine.Y < lead[0].StartBaseLine.Y - .01D && letter.StartBaseLine.Y > noteBaseline &&
            Math.Abs(letter.PointSize - lead[0].PointSize) < .001D).ToArray();
        Assert.Equal("TAIL", string.Concat(tail.Select(letter => letter.Value)));
        Assert.True(lead[0].StartBaseLine.Y > tail[0].StartBaseLine.Y);
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.Equal(1, actual.NumberOfPages);
            var rows = new[] {
                new FootnoteRuleBand(50.064, 363.460, 144.050, .600),
                new FootnoteRuleBand(146.090, 377.260, 132.050, .600),
                new FootnoteRuleBand(146.090, 391.060, 132.050, .600),
                new FootnoteRuleBand(146.090, 404.856, 132.050, .624),
                new FootnoteRuleBand(180.790, 418.680, 97.344, .600)
            };
            AssertFootnoteRuleBands(page, reference.GetPage(1), rows);
            // The mark consumes the remaining indented line; TAIL moves to the
            // next line. The alignment variants retain that same source frame.
            FootnoteRuleBand mixed = FindFootnoteRuleBand(page, rows[4]);
            Assert.True(mixed.X >= lead[lead.Length - 1].EndBaseLine.X - .15D);
            Assert.InRange(Math.Abs(tail[0].StartBaseLine.X - 146D), 0D, .15D);
            AssertPageFootnoteNativeText(page, reference.GetPage(1), "SHORT NOTE BODY",
                PageFootnoteNativeAnchor.PhysicalBottom);
        });
    }

    [Fact]
    public void SaveAsPdf_FootnoteRulePaintRetainsSourceRunMetrics() {
        PageFootnoteArtifact bytes = ConvertFootnoteRuleFixture("paint-font-size-style", mixedFonts: true);
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.Equal(1, actual.NumberOfPages);
            var rows = new[] {
                new FootnoteRuleBand(50.064, 314.236, 144.050, .504), // 8.5pt run/mark
                new FootnoteRuleBand(50.064, 327.340, 144.050, .720), // 13.5pt run/mark
                new FootnoteRuleBand(50.064, 341.860, 144.050, .600), // 12pt run / 6pt mark
                new FootnoteRuleBand(50.064, 351.580, 144.050, .240), // 6pt run / 12pt mark
                new FootnoteRuleBand(50.064, 362.620, 144.050, .600),
                new FootnoteRuleBand(50.064, 376.420, 144.050, .600), // bold
                new FootnoteRuleBand(50.064, 390.220, 144.050, .600), // principal underlined rule
                new FootnoteRuleBand(50.064, 401.136, 144.050, .360), // superscript
                new FootnoteRuleBand(50.064, 417.840, 144.050, .600, 0xC03040),
                new FootnoteRuleBand(50.064, 431.880, 144.050, .840)  // Calibri
            };
            AssertFootnoteRuleBands(actual.GetPage(1), reference.GetPage(1), rows);
            // Native drawing order puts this additional underline before the
            // principal band. Select its authored row/position, never path order.
            AssertFootnoteRuleBands(actual.GetPage(1), reference.GetPage(1),
                new[] { new FootnoteRuleBand(50.064, 394.660, 144.050, .840) }, "additional underline (source paragraph 6)");
            AssertPageFootnoteNativeText(actual.GetPage(1), reference.GetPage(1), "SHORT NOTE BODY",
                PageFootnoteNativeAnchor.PhysicalBottom);
        });
    }

    [Theory]
    [InlineData("continuation-notice-compat14", false)]
    [InlineData("continuation-notice-compat15", true)]
    public void SaveAsPdf_FootnoteRulesAndNoticesKeepNativeContinuationPartition(string id, bool ordinaryOnContinuation) {
        PageFootnoteArtifact bytes = ConvertFootnoteRuleFixture(id);
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        string notice = ordinaryOnContinuation ? "NOTICE MODE 15 CONTINUED ON NEXT PAGE" : "NOTICE MODE 14 CONTINUED ON NEXT PAGE";
        Assert.True(actual.NumberOfPages > 1);
        foreach (Page page in actual.GetPages()) {
            string expected = page.Number == 1 || ordinaryOnContinuation ? "START SEPARATOR" : "CONTINUATION SEPARATOR";
            string excluded = page.Number == 1 || ordinaryOnContinuation ? "CONTINUATION SEPARATOR" : "START SEPARATOR";
            PageFootnoteToken(page, expected);
            Assert.DoesNotContain(PageFootnoteCompact(excluded), PageFootnoteCompact(page.Text), StringComparison.Ordinal);
            if (page.Number < actual.NumberOfPages) PageFootnoteToken(page, notice);
            else Assert.DoesNotContain(PageFootnoteCompact(notice), PageFootnoteCompact(page.Text), StringComparison.Ordinal);
        }
        AssertPageFootnoteNativeProfile(bytes, () => {
            AssertFootnoteRulePartition(actual, reference, new[] { 40, 43, 29 });
            for (int page = 1; page <= 3; page++) {
                string label = page == 1 || ordinaryOnContinuation ? "START SEPARATOR" : "CONTINUATION SEPARATOR";
                AssertFootnoteRuleTextPosition(actual.GetPage(page), reference.GetPage(page), label);
                if (page < 3) AssertFootnoteRuleTextPosition(actual.GetPage(page), reference.GetPage(page), notice);
                double x = page == 1 || ordinaryOnContinuation ? 170.090 : 222.070;
                double width = page == 1 || ordinaryOnContinuation ? 144.050 : 128.090;
                double top = page == 1 ? 94.580 : page == 2 ? 58.580 : ordinaryOnContinuation ? 192.410 : 178.610;
                AssertFootnoteRuleBands(actual.GetPage(page), reference.GetPage(page),
                    new[] { new FootnoteRuleBand(x, top, width, .600) });
            }
        });
    }

    [Fact]
    public void SaveAsPdf_FootnoteRuleContinuesOldNoteBeforeTheNextAuthoredReferencePage() {
        PageFootnoteArtifact bytes = ConvertFootnoteRuleFixture("continuation-plus-new-note");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        int newReferencePage = PageFootnoteTokenPage(actual, "NEW NOTE REFERENCE");
        Assert.Equal(newReferencePage, PageFootnoteTokenPage(actual, "NEW NOTE BODY ON SAME CONTINUED PAGE"));
        Assert.Equal(newReferencePage, PageFootnoteTokenPage(actual, "FOLLOW BOTH NOTES"));
        Assert.True(newReferencePage > PageFootnoteTokenPage(actual, "NOTE 112"));
        foreach (Page page in actual.GetPages()) {
            PageFootnoteToken(page, "START SEPARATOR");
            Assert.DoesNotContain("CONTINUATIONSEPARATOR", PageFootnoteCompact(page.Text), StringComparison.Ordinal);
        }
        AssertPageFootnoteNativeProfile(bytes, () => {
            AssertFootnoteRulePartition(actual, reference, new[] { 41, 45, 26, 0 });
            Assert.Equal(4, newReferencePage);
            double[] tops = { 96.380, 60.380, 216.410, 432.480 };
            for (int page = 1; page <= 4; page++) {
                AssertFootnoteRuleBands(actual.GetPage(page), reference.GetPage(page),
                    new[] { new FootnoteRuleBand(170.090, tops[page - 1], 144.050, .600) });
                AssertFootnoteRuleTextPosition(actual.GetPage(page), reference.GetPage(page), "START SEPARATOR");
            }
            AssertFootnoteRuleTextPosition(actual.GetPage(4), reference.GetPage(4), "NEW NOTE REFERENCE");
            AssertPageFootnoteNativeText(actual.GetPage(4), reference.GetPage(4), "NEW NOTE BODY ON SAME CONTINUED PAGE",
                PageFootnoteNativeAnchor.PhysicalBottom);
        });
    }

    private PageFootnoteArtifact ConvertFootnoteRuleFixture(string id, bool mixedFonts = false,
        bool typedOrdinaryNoteInventory = false) {
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents", "WordPageFootnoteSeparators");
        string inputPath = Path.Combine(folder, id + ".docx");
        byte[] input = File.ReadAllBytes(inputPath);
        byte[] reference = File.ReadAllBytes(Path.Combine(folder, id + ".microsoft-word.pdf"));
        Assert.Equal(FootnoteRuleFixtureHashes[id].Input, PageFootnoteSha256(input));
        Assert.Equal(FootnoteRuleFixtureHashes[id].Reference, PageFootnoteSha256(reference));
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string[] before = { main.Document.OuterXml, main.FootnotesPart!.Footnotes!.OuterXml,
            main.StyleDefinitionsPart!.Styles!.OuterXml, main.DocumentSettingsPart!.Settings!.OuterXml,
            string.Join("\n", main.FooterParts.Select(part => part.Footer!.OuterXml)) };
        string ordinary = string.Join(" ", main.FootnotesPart.Footnotes.Elements<W.Footnote>()
            .Where(note => note.Type == null || note.Type.Value == W.FootnoteEndnoteValues.Normal).Select(note => note.InnerText));
        string[] tokens = Regex.Matches(ordinary, @"NOTE [0-9]{3}").Cast<Match>().Select(match => match.Value).ToArray();
        if (typedOrdinaryNoteInventory) AssertNoticeFixtureSourceProfile(document);
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        Assert.Equal(before, new[] { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml,
            main.StyleDefinitionsPart.Styles.OuterXml, main.DocumentSettingsPart.Settings.OuterXml,
            string.Join("\n", main.FooterParts.Select(part => part.Footer!.OuterXml)) });
        Assert.Empty(document.ValidateDocument());
        Assert.Equal(input, File.ReadAllBytes(inputPath));
        Assert.All(result.SourceConversionReports, report => Assert.False(report.HasLoss));
        var qualification = mixedFonts ? QualifyFootnoteRulePaintFonts(id, actual, result.Warnings) :
            QualifyPageFootnoteNativeArial(id, actual, result.Warnings);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
        if (qualification.Qualified) Assert.DoesNotContain(result.Warnings, warning => warning.Code is
            "NativeFootnoteSeparatorMarkOmitted" or "NativeFootnoteSeparatorDecorationOmitted");
        foreach (PdfConversionWarning warning in result.Warnings.Where(warning => warning.LossKind != OfficeConversionLossKind.None)) {
            bool font = qualification.Substitutions.Contains(warning) && warning.LossKind == OfficeConversionLossKind.Approximation;
            bool unavailableMetrics = !qualification.Qualified && (warning.Code is
                "NativeFootnoteSeparatorMarkOmitted" or "NativeFootnoteSeparatorDecorationOmitted") &&
                warning.Source == "footnote separator" && warning.LossKind == OfficeConversionLossKind.Omission;
            Assert.True(font || unavailableMetrics, $"Unexpected conversion loss: {warning}");
            _output.WriteLine("{0}: conversion loss: {1}", id, warning);
        }
        using var rendered = PdfPigDocument.Open(actual);
        if (typedOrdinaryNoteInventory) {
            NoticeNoteInventory inventory = ReadNoticeNoteInventory(rendered);
            Assert.Equal(PageFootnoteCompact(ordinary), inventory.Text);
            Assert.Equal(tokens.Select(PageFootnoteCompact), Regex.Matches(inventory.Text, @"NOTE[0-9]{3}")
                .Cast<Match>().Select(match => match.Value));
            foreach (string token in tokens) inventory.Token(token);
        } else if (tokens.Length == 0) {
            Assert.False(string.IsNullOrWhiteSpace(ordinary));
            AssertPageFootnoteTokenOnce(rendered, ordinary.Trim());
        }
        else {
            Assert.Equal(Enumerable.Range(1, 112).Select(number => $"NOTE {number:000}"), tokens);
            AssertFootnoteRuleTokenOrder(rendered, tokens);
        }
        return new PageFootnoteArtifact(actual, reference, ReadPageFootnoteFrame(input), qualification.Qualified, qualification.Qualification);
    }

    private (bool Qualified, string Qualification, PdfConversionWarning[] Substitutions) QualifyFootnoteRulePaintFonts(
        string id, byte[] actual, IReadOnlyList<PdfConversionWarning> warnings) {
        PdfConversionWarning[] substitutions = warnings.Where(warning =>
            warning.Code is "NativeFontFamilySubstituted" or "font-family-substitution").ToArray();
        foreach (PdfConversionWarning warning in substitutions) {
            Assert.Contains(warning.Details["fontFamily"], new[] { "Arial", "Calibri" });
            Assert.Equal("FontUnavailable", warning.Details["substitutionReason"]);
            _output.WriteLine("{0}: font degradation: {1}", id, warning);
        }
        bool arialAvailable = PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfEmbeddedFontFamily? arial);
        bool calibriAvailable = PdfEmbeddedFontFamily.TryFromSystem("Calibri", out PdfEmbeddedFontFamily? calibri);
        if (!arialAvailable) Assert.Contains(substitutions, warning => warning.Details["fontFamily"] == "Arial");
        if (!calibriAvailable) Assert.Contains(substitutions, warning => warning.Details["fontFamily"] == "Calibri");
        string regular = arialAvailable ? PageFootnoteSha256(arial!.Regular) : "unavailable";
        string bold = arial?.Bold is byte[] arialBold ? PageFootnoteSha256(arialBold) : "unavailable";
        string calibriRegular = calibriAvailable ? PageFootnoteSha256(calibri!.Regular) : "unavailable";
        using var rendered = PdfPigDocument.Open(actual);
        string[] faces = rendered.GetPages().SelectMany(page => page.Letters).Select(letter => letter.FontName ?? string.Empty)
            .Distinct(StringComparer.Ordinal).ToArray();
        bool declaredFamilies = faces.All(face => face.IndexOf("Arial", StringComparison.OrdinalIgnoreCase) >= 0 ||
            face.IndexOf("Calibri", StringComparison.OrdinalIgnoreCase) >= 0);
        if (!declaredFamilies) Assert.NotEmpty(substitutions);
        bool qualified = regular == NativeArialRegularHash && bold == NativeArialBoldHash &&
            calibriRegular == NativeCalibriRegularHash && declaredFamilies && substitutions.Length == 0;
        string profile = $"Arial regular={regular}, bold={bold}; Calibri regular={calibriRegular}; faces={string.Join(", ", faces)}; substitutions={substitutions.Length}";
        _output.WriteLine("{0}: native paint profile {1}: {2}", id, qualified ? "qualified" : "not qualified", profile);
        return (qualified, profile, substitutions);
    }

    private static void AssertFootnoteRuleTokenOrder(PdfPigDocument pdf, string[] tokens) {
        int previousPage = 0;
        double previousBaseline = double.PositiveInfinity;
        double previousX = 0D;
        foreach (string token in tokens) {
            AssertPageFootnoteTokenOnce(pdf, token);
            int page = PageFootnoteTokenPage(pdf, token);
            Letter first = PageFootnoteToken(pdf.GetPage(page), token)[0];
            Assert.True(page >= previousPage);
            if (page == previousPage) Assert.True(first.StartBaseLine.Y < previousBaseline - .01D ||
                Math.Abs(first.StartBaseLine.Y - previousBaseline) <= .01D && first.StartBaseLine.X > previousX,
                $"Source note token '{token}' precedes its previous token in the rendered flow.");
            previousPage = page;
            previousBaseline = first.StartBaseLine.Y;
            previousX = first.StartBaseLine.X;
        }
    }

    private static void AssertFootnoteRulePartition(PdfPigDocument actual, PdfPigDocument reference, int[] counts) {
        Assert.Equal(counts.Length, reference.NumberOfPages);
        Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
        int number = 1;
        for (int page = 1; page <= counts.Length; page++) {
            for (int count = 0; count < counts[page - 1]; count++, number++) {
                string token = $"NOTE {number:000}";
                Assert.Equal(page, PageFootnoteTokenPage(reference, token));
                Assert.Equal(page, PageFootnoteTokenPage(actual, token));
            }
        }
        Assert.Equal(113, number);
    }

    private static void AssertFootnoteRuleTextPosition(Page actual, Page reference, string token) {
        Letter observed = PageFootnoteToken(actual, token)[0];
        Letter expected = PageFootnoteToken(reference, token)[0];
        Assert.InRange(Math.Abs(observed.StartBaseLine.X - expected.StartBaseLine.X), 0D, .20D);
        Assert.InRange(Math.Abs(((double)actual.Height - observed.StartBaseLine.Y) -
            ((double)reference.Height - expected.StartBaseLine.Y)), 0D, .20D);
    }

    private static void AssertFootnoteRuleBands(Page actual, Page reference, FootnoteRuleBand[] expectedRows,
        string role = "principal") {
        for (int index = 0; index < expectedRows.Length; index++) {
            FootnoteRuleBand row = expectedRows[index];
            string identity = role + " row " + index;
            FootnoteRuleBand native = FindFootnoteRuleBand(reference, row, "reference " + identity);
            FootnoteRuleBand observed = FindFootnoteRuleBand(actual, row, "actual " + identity);
            Assert.InRange(Math.Abs(native.X - row.X), 0D, .001D);
            Assert.InRange(Math.Abs(native.Top - row.Top), 0D, .001D);
            Assert.InRange(Math.Abs(native.Width - row.Width), 0D, .001D);
            Assert.InRange(Math.Abs(native.Thickness - row.Thickness), 0D, .001D);
            // Four source twips cover Word's page/frame quantization. The much
            // smaller thickness budget admits its .024pt vector serialization
            // variation without hiding a missing 600-dpi device pixel (.12pt).
            Assert.InRange(Math.Abs(observed.X - native.X), 0D, .20D);
            Assert.InRange(Math.Abs(observed.Top - native.Top), 0D, .20D);
            Assert.InRange(Math.Abs(observed.Width - native.Width), 0D, .20D);
            Assert.InRange(Math.Abs(observed.Thickness - native.Thickness), 0D, .04D);
        }
    }

    private static FootnoteRuleBand FindFootnoteRuleBand(Page page, FootnoteRuleBand expected, string identity = "actual mixed row") {
        var filled = page.Paths.Select((path, index) => (Path: path, Index: index, Bounds: path.GetBoundingRectangle()))
            .Where(item => item.Path.IsFilled).ToArray();
        var candidates = filled
            .Where(item => item.Bounds is { } bounds && bounds.Width > 20D && bounds.Height > .1D && bounds.Height < 1.5D &&
                Math.Abs(bounds.Left - expected.X) < 2D && Math.Abs(bounds.Right - expected.X - expected.Width) < 2D &&
                Math.Abs((double)page.Height - bounds.Top - expected.Top) < 2D).ToArray();
        string[] horizontal = filled.Where(item => item.Bounds is { Width: > 20D, Height: >= 0D and < 10D })
            .Select(item => {
                var bounds = item.Bounds!.Value;
                var rgb = item.Path.FillColor.ToRGBValues();
                return FormattableString.Invariant($"path={item.Index}; x={bounds.Left:F6}; top={(double)page.Height - bounds.Top:F6}; right={bounds.Right:F6}; width={bounds.Width:F6}; height={bounds.Height:F6}; rgb=({rgb.Item1:F6},{rgb.Item2:F6},{rgb.Item3:F6})");
            }).ToArray();
        Assert.True(candidates.Length == 1, FormattableString.Invariant(
            $"{identity}; page={page.Number}; pageSize={page.Width:F6}x{page.Height:F6}; expected x={expected.X:F6}, top={expected.Top:F6}, width={expected.Width:F6}, thickness={expected.Thickness:F6}, color=#{expected.Color:X6}; matches={candidates.Length}; all filled horizontal paths: {string.Join(" | ", horizontal)}"));
        var selected = candidates[0];
        var rectangle = selected.Bounds!.Value;
        var color = selected.Path.FillColor.ToRGBValues();
        Assert.InRange(Math.Abs(color.Item1 - ((expected.Color >> 16) & 255) / 255D), 0D, .001D);
        Assert.InRange(Math.Abs(color.Item2 - ((expected.Color >> 8) & 255) / 255D), 0D, .001D);
        Assert.InRange(Math.Abs(color.Item3 - (expected.Color & 255) / 255D), 0D, .001D);
        return new FootnoteRuleBand(rectangle.Left, (double)page.Height - rectangle.Top, rectangle.Width, rectangle.Height, expected.Color);
    }

    private sealed record FootnoteRuleBand(double X, double Top, double Width, double Thickness, int Color = 0);
}
