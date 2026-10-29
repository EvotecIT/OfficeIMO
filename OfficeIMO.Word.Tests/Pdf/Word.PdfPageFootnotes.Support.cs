using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public partial class WordPdfPageFootnoteTests {
    private readonly ITestOutputHelper _output;
    private const string NativeArialRegularHash = "b3658eadae55e682b5f69eb64c439c1ecc8f196c0bb8d4756d145d13bc86476a";
    private const string NativeArialBoldHash = "e8f4e3baf6cc35fed6fcce3a540e8b39e8f6cda1d22a28f2ec8f526fef7a43f5";
    private const string SeparatorWarningCode = "native-footnote-separator-mark-omitted";

    public WordPdfPageFootnoteTests(ITestOutputHelper output) => _output = output;

    // These pairs are owned synthetic inputs and read-only native Word exports,
    // not round trips through the converter under test. The adjacent provenance
    // records the producer, font profile, inert-package checks and schema result.
    private static readonly IReadOnlyDictionary<string, (string Input, string Reference)> PageFootnoteFixtureHashes =
        new Dictionary<string, (string Input, string Reference)>(StringComparer.Ordinal) {
            ["footnote-doc-beneathText-section-pageBottom"] = ("da245b9e4be283e0670ee01fd70de81673013a1a24c4005d60377eed0092af47", "230f1be2ade0fd98577cb2e579bf32ee08eacd7b36830c88e6fe407c96abf6af"),
            ["footnote-doc-pageBottom-section-beneathText"] = ("1943f4ffefe1852225dc5205c2f0865a8f1b348c604a9c1ee8b2a096c56d1f95", "235785712973ca649219b49de3fb057db1994184669456886c2dce67c39f1cbe"),
            ["fallback-doc-beneathText-section-absent"] = ("6b1db5bb1a2d988125861c5cd7091cc5063a18875cf29c2fc06425732682ac28", "de6afed0848088d6a2a77c05498fb2d55558e79aeb1ce2e38f0c64edc82bd907"),
            ["fallback-doc-beneathText-sectionFootnotePr-absent"] = ("6aec55da4a1b357908024eef8594f2ecd8f6ab7b27ee2fbb1eb374b8fc45072c", "4eaa2db03a80302d7eb7f29cf7592365751009d159ae1d29fab0ac903d6ed035"),
            ["fallback-doc-absent-sectionFootnotePr-absent"] = ("e405fb99a1de25ab4e1cad86eb31985c61d61367eec522aadf3b72660fb11a27", "15421f69a48bcbc9c865749c0db0573257b4fcd8e3b915c7c947194b0090850e"),
            ["richnote-long-rich-continuation"] = ("43393e76b7de826bb2d1c02cc71b8fb183e6dee373d08bd8176d3784a487e263", "ba436fbce97297b43a6d7e1fb246f647269cfefbd60eafa5a44e221ade0bee76"),
            ["richnote-near-break-fill24"] = ("772fbcb7476241b38c262d9bb5c892e03593867147236c2f5415d58904796340", "a447c096344477c6a179bbfec4f88b673e6b2f2dc583bff5fe19f162a487f4f0"),
            ["richnote-near-break-fill25"] = ("dbe86fabf5f45eb4179f7f403659400885623eb3ee2dc6a1ddfc622620dad818", "ecd6bd16eeea17251c24e7dd96fea5119726e553a643605c042bcdbdb57ed87b"),
            ["richnote-near-break-fill27"] = ("c71caf4de640db7e38023853e70047029ec11a24167c145a3eefa7d247ce3467", "761d9268fd2a78c022a640353c6c009e7bca0ec2358cbb178b56ad7c37a095ff"),
            ["richnote-picture-internal-link"] = ("1f8a207c8dadf2137cba23de557eaf2978d99c7c31a49f67f488965fabb0b32d", "433b6d83e91764bd0dc0371a34e4f79f10b0c86ef6ad24c3fd72afe645cc0684"),
            ["richnote-table-late-fragment"] = ("d0a69539df472960122bfe0b0e858e503ae20075f1776c9bc9b3ecfdacca3d08", "d3cc5b13ad40164bfde0a0fcd89359467fb94d55f69d4b2d2509ff03cca710ac"),
            ["richnote-two-notes-one-page"] = ("dbb89559739775cd478ed37963fa6761acc3902de501b6789b8a9b259d10eadd", "0d2c38f894bd4b10e4f804c4bff0f0997c3605bf606706b923a247216fc6a16f")
        };

    private PageFootnoteArtifact ConvertPageFootnoteFixture(string id) {
        string inputPath = PageFootnoteFixturePath(id + ".docx");
        string referencePath = PageFootnoteFixturePath(id + ".microsoft-word.pdf");
        byte[] input = File.ReadAllBytes(inputPath);
        byte[] reference = File.ReadAllBytes(referencePath);
        Assert.Equal(PageFootnoteFixtureHashes[id].Input, PageFootnoteSha256(input));
        Assert.Equal(PageFootnoteFixtureHashes[id].Reference, PageFootnoteSha256(reference));
        PageFootnoteFrame frame = ReadPageFootnoteFrame(input);
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.ValidateDocument());
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        Assert.Equal("%PDF", Encoding.ASCII.GetString(actual, 0, 4));
        Assert.Equal(input, File.ReadAllBytes(inputPath));
        var qualification = QualifyPageFootnoteNativeArial(id, actual, result.Warnings);

        AssertPageFootnoteAcceptedLosses(id, result, qualification.Qualified, qualification.Substitutions);
        return new PageFootnoteArtifact(actual, reference, frame, qualification.Qualified, qualification.Qualification);
    }

    private void AssertPageFootnoteAcceptedLosses(string id, PdfDocumentConversionResult result,
        bool nativeArialQualified, PdfConversionWarning[] substitutions) {
        PdfConversionWarning[] separators = result.Warnings.Where(warning => warning.Code == SeparatorWarningCode).ToArray();
        if (nativeArialQualified) Assert.Empty(separators);
        else Assert.InRange(separators.Length, 0, 1);
        foreach (PdfConversionWarning warning in separators) {
            Assert.Equal("footnote separator", warning.Source);
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.Contains("selected font does not provide usable strike-band metrics", warning.Message, StringComparison.Ordinal);
            _output.WriteLine("{0}: separator font-metric limitation: {1}", id, warning.Message);
        }
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "native-footnote-placement-approximated");
        Assert.All(result.SourceConversionReports, report => Assert.False(report.HasLoss));
        Assert.All(substitutions, warning => Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind));
        foreach (PdfConversionWarning warning in result.Warnings.Where(warning => warning.LossKind != OfficeConversionLossKind.None)) {
            bool knownSeparator = separators.Contains(warning);
            bool knownFont = substitutions.Contains(warning);
            Assert.True(knownSeparator || knownFont, $"Unexpected conversion loss: {warning}");
        }
        Assert.Equal(separators.Length > 0 || substitutions.Length > 0, result.HasLoss);
    }

    private (bool Qualified, string Qualification, PdfConversionWarning[] Substitutions) QualifyPageFootnoteNativeArial(
        string id, byte[] actual, IReadOnlyList<PdfConversionWarning> warnings,
        IReadOnlyCollection<string>? declaredSourceFontFamilies = null) {
        using var rendered = PdfPigDocument.Open(actual);
        string[] faces = rendered.GetPages().SelectMany(page => page.Letters).Select(letter => letter.FontName ?? string.Empty)
            .Distinct(StringComparer.Ordinal).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        PdfConversionWarning[] substitutions = warnings.Where(warning =>
            warning.Code == "font-family-substitution").ToArray();
        IReadOnlyCollection<string> sourceFamilies = declaredSourceFontFamilies ?? new[] { "Arial" };
        foreach (PdfConversionWarning warning in substitutions) {
            _output.WriteLine("{0}: font degradation: code={1}; source={2}; loss={3}; {4}; details={5}",
                id, warning.Code, warning.Source, warning.LossKind, warning.Message,
                string.Join(", ", warning.Details.Select(detail => detail.Key + "=" + detail.Value)));
        }
        foreach (PdfConversionWarning warning in substitutions) {
            string sourceFamily = warning.Details["fontFamily"];
            Assert.Contains(sourceFamily, sourceFamilies);
            // The shared resolver and the Word adapter use the same catalog code,
            // with locations identifying their distinct source boundaries.
            Assert.Contains(warning.Source, new[] { "font:" + sourceFamily, "word:font[" + sourceFamily + "]" });
            Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
            Assert.Equal("FontUnavailable", warning.Details["substitutionReason"]);
            Assert.True(warning.Details.ContainsKey("fallbackSlot") || warning.Details.ContainsKey("resolvedFontFamily"));
            if (warning.Source == "font:" + sourceFamily) {
                Assert.Equal("OfficeDocumentFonts", warning.Details["substitutionPolicy"]);
                Assert.Equal("NotPerformed", warning.Details["fontMetricsVerification"]);
            } else {
                Assert.Equal("Helvetica", warning.Details["fallbackSlot"]);
            }
        }
        bool declaredArial = faces.All(face => face.IndexOf("Arial", StringComparison.OrdinalIgnoreCase) >= 0);
        bool familyAvailable = PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfEmbeddedFontFamily? family);
        if (!declaredArial || !familyAvailable) Assert.NotEmpty(substitutions);
        string regularHash = familyAvailable ? PageFootnoteSha256(family!.Regular) : "unavailable";
        byte[]? bold = family?.Bold;
        string boldHash = bold == null ? "unavailable" : PageFootnoteSha256(bold);
        bool qualified = regularHash == NativeArialRegularHash && boldHash == NativeArialBoldHash &&
            declaredArial && substitutions.Length == 0;
        string qualification = $"Arial7.06 regular={regularHash}, bold={boldHash}; faces={string.Join(", ", faces)}; substitutions={substitutions.Length}";
        _output.WriteLine("{0}: native profile {1}: {2}", id, qualified ? "qualified" : "not qualified", qualification);
        return (qualified, qualification, substitutions);
    }

    private static string PageFootnoteFixturePath(string file) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "WordPageFootnotes", file);

    private void AssertPageFootnoteNativeProfile(PageFootnoteArtifact artifact, Action? assertions = null) {
        // The strict coordinator-owned Windows gate requires the exact source
        // font programs. Portable assertions have already run before this call.
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_REQUIRE_NATIVE_ARIAL") == "1")
            Assert.True(artifact.NativeArialQualified, "Required native profile unavailable: " + artifact.Qualification);
        if (artifact.NativeArialQualified && assertions != null) {
            _output.WriteLine("Native geometry: qualified; executing the unchanged native budgets.");
            assertions();
            _output.WriteLine("Native geometry: passed.");
        } else if (artifact.NativeArialQualified)
            _output.WriteLine("Native profile: qualified; this source-order contract has no additional native coordinate budget.");
        else _output.WriteLine("Native geometry: not evaluated; portable contract assertions passed. " + artifact.Qualification);
    }

    private static PageFootnoteFrame ReadPageFootnoteFrame(byte[] input) {
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(input), false);
        var main = package.MainDocumentPart!;
        W.SectionProperties section = Assert.Single(main.Document.Body!.Elements<W.SectionProperties>());
        W.PageMargin margin = section.GetFirstChild<W.PageMargin>()!;
        W.Style noteStyle = Assert.Single(main.StyleDefinitionsPart!.Styles!.Elements<W.Style>(), style => style.StyleId?.Value == "FootnoteText");
        W.SpacingBetweenLines spacing = noteStyle.StyleParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal(W.LineSpacingRuleValues.Exact, spacing.LineRule!.Value);
        bool hasSeparatorMark = main.FootnotesPart!.Footnotes!.Descendants().Any(element =>
            element.LocalName is "separator" or "continuationSeparator");
        return new PageFootnoteFrame(margin.Top!.Value / 20D, margin.Bottom!.Value / 20D,
            double.Parse(spacing.Line!.Value!, CultureInfo.InvariantCulture) / 20D, hasSeparatorMark);
    }

    private sealed record PageFootnoteFrame(double MarginTop, double MarginBottom, double NoteLineHeight, bool HasSeparatorMark);
    private sealed record PageFootnoteArtifact(byte[] Actual, byte[] Reference, PageFootnoteFrame Frame,
        bool NativeArialQualified, string Qualification);

    private static string PageFootnoteSha256(byte[] bytes) {
        using SHA256 algorithm = SHA256.Create();
        return BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }

    private static string PageFootnoteCompact(string text) => Regex.Replace(text, @"\s+", string.Empty);

    private static Letter[] PageFootnoteToken(Page page, string token) {
        // Keep a glyph map while ignoring extraction-only spaces. This permits
        // separate PDF text operators without requiring one span per source run.
        var letters = new List<Letter>();
        var text = new StringBuilder();
        foreach (Letter letter in page.Letters) {
            foreach (char character in letter.Value) {
                if (char.IsWhiteSpace(character)) continue;
                text.Append(character);
                letters.Add(letter);
            }
        }
        string compact = text.ToString();
        string needle = PageFootnoteCompact(token);
        int start = compact.IndexOf(needle, StringComparison.Ordinal);
        Assert.True(start >= 0, $"Expected '{token}' on page {page.Number}.");
        Assert.Equal(-1, compact.IndexOf(needle, start + needle.Length, StringComparison.Ordinal));
        return letters.Skip(start).Take(needle.Length).ToArray();
    }

    private static int PageFootnoteTokenPage(PdfPigDocument pdf, string token) =>
        Assert.Single(pdf.GetPages(), page => PageFootnoteCompact(page.Text)
            .Contains(PageFootnoteCompact(token), StringComparison.Ordinal)).Number;

    private static void AssertPageFootnoteTokenOnce(PdfPigDocument pdf, string token) {
        string compact = string.Concat(pdf.GetPages().Select(page => PageFootnoteCompact(page.Text)));
        Assert.Single(Regex.Matches(compact, Regex.Escape(PageFootnoteCompact(token))));
    }

    private static void AssertPageFootnotePortableStyle(Page page, string token, bool bold = false, int color = 0) {
        foreach (Letter letter in PageFootnoteToken(page, token)) {
            Assert.Equal(10D, letter.PointSize, 3);
            Assert.Equal(bold, letter.FontName.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0);
            var rgb = letter.Color.ToRGBValues();
            Assert.Equal(((color >> 16) & 255) / 255D, rgb.Item1, 3);
            Assert.Equal(((color >> 8) & 255) / 255D, rgb.Item2, 3);
            Assert.Equal((color & 255) / 255D, rgb.Item3, 3);
        }
    }

    private static void AssertPageFootnoteBottomLine(Page page, string lastNoteToken, PageFootnoteFrame frame) {
        double baseline = PageFootnoteToken(page, lastNoteToken)[0].StartBaseLine.Y;
        // This checks the authored final exact-height line frame, independently
        // of a font's particular baseline inside that frame.
        Assert.InRange(baseline, frame.MarginBottom, frame.MarginBottom + frame.NoteLineHeight);
        var footer = PageFootnoteToken(page, "FOOTER");
        Assert.True(PageFootnoteToken(page, lastNoteToken).Min(letter => letter.BoundingBox.Bottom) >
            footer.Max(letter => letter.BoundingBox.Top));
    }

    private static void AssertPageFootnoteBodyDoesNotOverlapNotes(PdfPigDocument pdf,
        IEnumerable<string> bodyTokens, IEnumerable<string> noteTokens) {
        var notes = noteTokens.GroupBy(token => PageFootnoteTokenPage(pdf, token)).ToDictionary(group => group.Key,
            group => group.SelectMany(token => PageFootnoteToken(pdf.GetPage(group.Key), token)).ToArray());
        foreach (string token in bodyTokens) {
            int pageNumber = PageFootnoteTokenPage(pdf, token);
            if (!notes.TryGetValue(pageNumber, out Letter[]? noteLetters)) continue;
            Letter[] body = PageFootnoteToken(pdf.GetPage(pageNumber), token);
            Assert.True(body.Min(letter => letter.BoundingBox.Bottom) >= noteLetters.Max(letter => letter.BoundingBox.Top),
                $"Body '{token}' overlaps the note text frame on page {pageNumber}.");
        }
    }

    private static void AssertPageFootnoteParagraphOrder(PdfPigDocument pdf, string[] tokens, PageFootnoteFrame frame) {
        int priorPage = 0;
        double priorBaseline = double.PositiveInfinity;
        foreach (string token in tokens) {
            AssertPageFootnoteTokenOnce(pdf, token);
            int pageNumber = PageFootnoteTokenPage(pdf, token);
            Page page = pdf.GetPage(pageNumber);
            double baseline = PageFootnoteToken(page, token)[0].StartBaseLine.Y;
            Assert.InRange(baseline, frame.MarginBottom, (double)page.Height - frame.MarginTop);
            Assert.True(pageNumber >= priorPage);
            if (pageNumber == priorPage) Assert.Equal(frame.NoteLineHeight, priorBaseline - baseline, 3);
            priorPage = pageNumber;
            priorBaseline = baseline;
        }
        foreach (var group in tokens.GroupBy(token => PageFootnoteTokenPage(pdf, token)))
            AssertPageFootnoteBottomLine(pdf.GetPage(group.Key), group.Last(), frame);
    }

    private enum PageFootnoteNativeAnchor { PhysicalBottom, PrecedingBodyBaseline }

    private void AssertPageFootnoteNativeText(Page actual, Page reference, string token,
        PageFootnoteNativeAnchor anchor, string? precedingBodyToken = null) {
        Letter[] expected = PageFootnoteToken(reference, token);
        Letter[] observed = PageFootnoteToken(actual, token);
        Assert.Equal(expected.Length, observed.Length);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Contains("Arial", observed[index].FontName, StringComparison.OrdinalIgnoreCase);
            // Native Word rounds the authored 10pt face to 9.96pt in PDF;
            // preserve that source size without requiring identical rounding.
            Assert.InRange(Math.Abs(expected[index].PointSize - observed[index].PointSize), 0D, .05D);
            Assert.Equal(expected[index].FontName.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0,
                observed[index].FontName.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0);
            var expectedColor = expected[index].Color.ToRGBValues();
            var observedColor = observed[index].Color.ToRGBValues();
            Assert.Equal(expectedColor.Item1, observedColor.Item1, 3);
            Assert.Equal(expectedColor.Item2, observedColor.Item2, 3);
            Assert.Equal(expectedColor.Item3, observedColor.Item3, 3);
        }
        double nativeBottomY = expected[0].StartBaseLine.Y;
        double actualBottomY = observed[0].StartBaseLine.Y;
        double nativeTopY = (double)reference.Height - nativeBottomY;
        double actualTopY = (double)actual.Height - actualBottomY;
        if (anchor == PageFootnoteNativeAnchor.PhysicalBottom) {
            // Page-bottom content retains its physical bottom anchor when the
            // native PDF quantizes the page height differently from the source.
            _output.WriteLine(FormattableString.Invariant(
                $"{token}: anchor=physical-bottom; nativeBottomY={nativeBottomY:F6}; actualBottomY={actualBottomY:F6}; nativeTopY={nativeTopY:F6}; actualTopY={actualTopY:F6}; nativePageHeight={reference.Height:F6}; actualPageHeight={actual.Height:F6}"));
            Assert.InRange(Math.Abs(nativeBottomY - actualBottomY), 0D, .15D);
        } else {
            Assert.False(string.IsNullOrWhiteSpace(precedingBodyToken));
            double nativeBodyY = PageFootnoteToken(reference, precedingBodyToken!)[0].StartBaseLine.Y;
            double actualBodyY = PageFootnoteToken(actual, precedingBodyToken!)[0].StartBaseLine.Y;
            double nativeGap = nativeBodyY - nativeBottomY;
            double actualGap = actualBodyY - actualBottomY;
            // Beneath-text placement belongs to the preceding body baseline;
            // retain raw absolute/body coordinates alongside the relative proof.
            _output.WriteLine(FormattableString.Invariant(
                $"{token}: anchor=body-baseline({precedingBodyToken}); nativeTopY={nativeTopY:F6}; actualTopY={actualTopY:F6}; nativeBodyTopY={(double)reference.Height - nativeBodyY:F6}; actualBodyTopY={(double)actual.Height - actualBodyY:F6}; nativeGap={nativeGap:F6}; actualGap={actualGap:F6}; rawAbsoluteDelta={actualTopY - nativeTopY:F6}"));
            Assert.InRange(Math.Abs(nativeGap - actualGap), 0D, .15D);
        }
        Assert.InRange(Math.Abs(expected[0].StartBaseLine.X - observed[0].StartBaseLine.X), 0D, .15D);
    }

    private static void AssertPageFootnoteInlineReference(Page page, string prefix, string suffix, string marker) {
        Letter[] before = PageFootnoteToken(page, prefix);
        Letter[] after = PageFootnoteToken(page, suffix);
        Letter reference = Assert.Single(page.Letters, letter => letter.Value == marker &&
            letter.StartBaseLine.Y > before[0].StartBaseLine.Y &&
            letter.StartBaseLine.Y - before[0].StartBaseLine.Y < before[0].PointSize);
        Assert.InRange(reference.StartBaseLine.X, before[before.Length - 1].EndBaseLine.X - .15D,
            after[0].StartBaseLine.X + .15D);
        Assert.True(reference.EndBaseLine.X <= after[0].StartBaseLine.X + .15D);
        Assert.True(reference.PointSize < before[0].PointSize);
        Assert.Equal(before[0].StartBaseLine.Y, after[0].StartBaseLine.Y, 3);
    }
}
