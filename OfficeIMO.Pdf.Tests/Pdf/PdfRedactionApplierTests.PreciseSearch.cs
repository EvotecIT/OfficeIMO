using OfficeIMO.Pdf;
using System.Text.RegularExpressions;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfRedactionApplierTests {
        [Fact]
        public void PreciseSearchJoinsTouchingTextShowsWithoutJoiningSeparateOccurrences() {
            PdfDocument source = PdfDocument.Load(BuildTextContentRedactionSource(
                "BT /F1 20 Tf 72 720 Td (Alpha ) Tj (secret) Tj ( account) Tj ( Omega secret account) Tj ET"));
            PdfRedactionPlan search = source.Redactions.Search(PreciseOptions().AddLiteral("secret account"));
            Assert.True(search.IsReviewable, DescribeFindings(search));
            Assert.Collection(search.Areas, area => Assert.Equal(1, area.PageNumber), area => Assert.Equal(1, area.PageNumber));
            PdfRedactionPlan selection = source.Redactions.Plan(new[] { search.Areas[0].WithLabel("Selected occurrence") });
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(selection);
            Assert.True(result.Evidence.IsVerified);
            string text = result.ToDocument().Read().Text;
            Assert.Equal(1, Regex.Matches(text, "secret account").Count);
            Assert.Contains("Alpha", text, StringComparison.Ordinal);
            Assert.Contains("Omega", text, StringComparison.Ordinal);
        }

        [Fact]
        public void RelabelingPreciseAreasPreservesTheRemovalContract() {
            byte[] bytes = BuildSingleTextObjectRedactionSource("(Alpha secret Omega) Tj");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionArea area = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("secret")).Areas);
            PdfRedactionPlan relabeled = source.Redactions.Plan(new[] { area.WithLabel("Reviewed reason") });
            Assert.Equal("Reviewed reason", Assert.Single(relabeled.Areas).Label);
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(relabeled);
            Assert.True(result.Evidence.IsVerified);
            string text = result.ToDocument().Read().Text;
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
            Assert.Contains("Alpha", text, StringComparison.Ordinal);
            Assert.Contains("Omega", text, StringComparison.Ordinal);
            Assert.Equal(bytes, source.ToBytes());
        }

        [Fact]
        public void PreciseSelectionBlocksIntersectingUnselectedHiddenLayerGlyphs() {
            const string content = "BT /F1 20 Tf 72 720 Td (secret) Tj ET " +
                "/OC /Layer BDC BT /F1 20 Tf 72 720 Td (private) Tj ET EMC";
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /OFF [6 0 R] >> >> >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /Properties << /Layer 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(content)),
                "6 0 obj\n<< /Type /OCG /Name (Hidden layer) >>\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            Assert.DoesNotContain("private", source.Reader.Text(), StringComparison.Ordinal);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "RedactionSearchUnselectedTextIntersection");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
            Assert.Equal(bytes, source.ToBytes());
        }

        [Theory]
        [InlineData("(WAAA) Tj")]
        [InlineData("[(WA) 277 (WA)] TJ")]
        public void PreciseSelectionBlocksCoincidentUnselectedGlyphOccurrences(string operation) {
            byte[] bytes = BuildTextContentRedactionSource("BT /F1 20 Tf -13.34 Tc 72 720 Td " + operation + " ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddRegex("^WA"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "RedactionSearchUnselectedTextIntersection");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
            Assert.Equal(bytes, source.ToBytes());
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        public void TextOnlySharingEvidenceExcludesPreservedVectorAndImageUnderlays(PdfRedactionTextSelection selection) {
            const string content = "q 0.2 0.3 0.4 rg 60 680 350 65 re f Q " +
                "q 350 0 0 65 60 680 cm /Im1 Do Q BT /F1 20 Tf 72 720 Td (Alpha secret Omega) Tj ET";
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Im1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(content)),
                "6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionSearchOptions options = PreciseOptions().AddLiteral("secret");
            options.TextSelection = selection;
            PdfRedactionPlan plan = source.Redactions.Search(options);
            Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.VectorPath);
            Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.ImagePlacement);

            PdfRedactionSharingResult result = source.Redactions.ApplyForSharing(plan,
                new PdfSanitizationOptions { ContentKindsToRemove = PdfSanitizationContentKind.UserMetadata });

            Assert.True(result.Redaction.IsVerified, result.Redaction.Summary);
            Assert.Empty(result.Redaction.ResidualMatches);
            Assert.All(result.Redaction.Items, item => Assert.Equal(PdfRedactionMatchKind.TextBlock, item.ReviewedMatch.Kind));
            string raw = PdfEncoding.Latin1GetString(result.ToBytes());
            Assert.Contains("60 680 350 65 re f", raw, StringComparison.Ordinal);
            Assert.Contains("/Subtype /Image", raw, StringComparison.Ordinal);
            string text = result.Sanitization.ToDocument().Read().Text;
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
            if (selection == PdfRedactionTextSelection.MatchedGlyphs) {
                Assert.Contains("Alpha", text, StringComparison.Ordinal);
                Assert.Contains("Omega", text, StringComparison.Ordinal);
            }
            Assert.Equal(bytes, source.ToBytes());
        }

        [Theory]
        [InlineData("(Alpha secret Omega) Tj")]
        [InlineData("[(Alpha ) -120 (secret) -120 ( Omega)] TJ")]
        [InlineData("(Alpha secret Omega) '")]
        [InlineData("0 0 (Alpha secret Omega) \"")]
        [InlineData("(Alpha   secret   Omega) Tj")]
        [InlineData("[(Alpha   ) -120 (secret) -120 (   Omega)] TJ")]
        public void PreciseSearchRetainsUnselectedEncodedGlyphsAndTheirPositions(string showOperation) {
            byte[] bytes = BuildSingleTextObjectRedactionSource(showOperation);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Single(plan.Areas);

            PdfDocument result = source.Redactions.Apply(plan);

            string text = result.Read().Text;
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
            Assert.Contains("Alpha", text, StringComparison.Ordinal);
            Assert.Contains("Omega", text, StringComparison.Ordinal);
            (string Text, double X, double Y, string Code, string Font)[] before = ReadGlyphEvidence(bytes).Where(glyph => glyph.X < Math.Round(plan.Areas[0].X, 5) || glyph.X >= Math.Round(plan.Areas[0].Right, 5)).ToArray();
            (string Text, double X, double Y, string Code, string Font)[] after = ReadGlyphEvidence(result.ToBytes()).ToArray();
            Assert.Equal(before, after);
            Assert.Equal(bytes, source.ToBytes());
        }

        [Fact]
        public void PreciseRegexSelectsEveryOccurrenceAndPreservesOtherText() {
            PdfDocument source = PdfDocument.Load(BuildSingleTextObjectRedactionSource("(Alpha secret1 Beta secret2 Omega) Tj"));
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddRegex("secret[0-9]"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Equal(2, plan.Areas.Count);
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified);
            string text = result.ToDocument().Read().Text;
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
            Assert.Contains("Alpha", text, StringComparison.Ordinal);
            Assert.Contains("Beta", text, StringComparison.Ordinal);
            Assert.Contains("Omega", text, StringComparison.Ordinal);
        }

        [Fact]
        public void PreciseLiteralMatchesAcrossTextShowsAndWrappedLines() {
            byte[] bytes = BuildTextContentRedactionSource("BT /F1 12 Tf 22 TL 72 720 Td (Alpha secret) Tj T* (account Omega) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret account"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Equal(2, plan.Areas.Count);
            PdfDocument result = source.Redactions.Apply(plan);
            string text = result.Read().Text;
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
            Assert.DoesNotContain("account", text, StringComparison.Ordinal);
            Assert.Contains("Alpha", text, StringComparison.Ordinal);
            Assert.Contains("Omega", text, StringComparison.Ordinal);
        }

        [Fact]
        public void PreciseSearchAndApplyPreservesNeighboursInsideNestedForms() {
            PdfDocument source = PdfDocument.Load(BuildNestedFormXObjectRedactionSource());
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            PdfDocument result = source.Redactions.Apply(plan);
            string text = result.Read().Text;
            Assert.Contains("Nested", text, StringComparison.Ordinal);
            Assert.Contains("account 123-45", text, StringComparison.Ordinal);
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("0 1 -1 0 200 220")]
        [InlineData("0 -1 1 0 200 620")]
        [InlineData("-1 0 0 -1 500 620")]
        public void PreciseSelectionPreservesGlyphOriginsUnderCardinalTextRotation(string matrix) {
            byte[] bytes = BuildTextContentRedactionSource("BT /F1 20 Tf " + matrix + " Tm (Alpha secret Omega) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            PdfDocument result = source.Redactions.Apply(plan);
            (string Text, double X, double Y, string Code, string Font)[] before = ReadGlyphEvidence(bytes).ToArray();
            Assert.Equal(before.Take(6).Concat(before.Skip(12)), ReadGlyphEvidence(result.ToBytes()));
            Assert.DoesNotContain("secret", result.Read().Text, StringComparison.Ordinal);
        }

        [Fact]
        public void PreciseSelectionKeepsWholeLigaturesAndRejectsPartialOnes() {
            byte[] bytes = BuildType0ToUnicodeLigatureRedactionSource();
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan partial = source.Redactions.Search(PreciseOptions().AddLiteral("f"));
            Assert.False(partial.IsReviewable);
            Assert.Contains(partial.Findings, finding => finding.Code == "RedactionSearchPartialGlyph");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(partial));

            PdfRedactionPlan whole = source.Redactions.Search(PreciseOptions().AddLiteral("fi"));
            Assert.True(whole.IsReviewable, DescribeFindings(whole));
            string text = source.Redactions.Apply(whole).Read().Text;
            Assert.DoesNotContain("fi", text, StringComparison.Ordinal);
            Assert.Contains("A", text, StringComparison.Ordinal);
            Assert.Contains("Z", text, StringComparison.Ordinal);
            Assert.Equal(bytes, source.ToBytes());
        }

        [Theory]
        [InlineData("/Span << /ActualText (Alpha secret Omega) >> BDC (painted substitute) Tj EMC")]
        [InlineData("q 60 700 m 410 700 l 100 740 l h W n BT /F1 20 Tf 72 720 Td (Alpha secret Omega) Tj ET Q")]
        public void PreciseSelectionBlocksAmbiguousOrClippedSourceText(string content) {
            byte[] bytes = content.StartsWith("q ", StringComparison.Ordinal)
                ? BuildTextContentRedactionSource(content) : BuildSingleTextObjectRedactionSource(content);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "RedactionSearchGlyphMappingUnsupported");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
            Assert.Equal(bytes, source.ToBytes());
        }

        [Theory]
        [InlineData("BT /F1 20 Tf -8 Tc 72 720 Td (AB) Tj ET", "A")]
        [InlineData("BT /F1 12 Tf 14 TL 72 720 Td (Alpha secret) Tj T* (account Omega) Tj ET", "secret account")]
        [InlineData("BT /F1 20 Tf 72 720 Td [(Alpha ) 120 (secret) -120 ( Omega)] TJ ET", "secret")]
        public void PreciseSelectionBlocksOverlappingUnselectedGlyphs(string content, string literal) {
            byte[] bytes = BuildTextContentRedactionSource(content);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral(literal));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "RedactionSearchUnselectedTextIntersection");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
        }

        [Fact]
        public void PreciseSelectionHonorsPagesCaseCancellationAndCandidateLimits() {
            PdfDocument source = PdfDocument.Create(pdf => {
                pdf.Page(page => page.Content(content => content.Text("Alpha SECRET Omega")));
                pdf.Page(page => page.Content(content => content.Text("Alpha SECRET Omega")));
            });
            PdfRedactionSearchOptions options = PreciseOptions().AddLiteral("secret");
            options.MaximumCandidates = 1;
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Search(options));
            options.PageNumbers.Add(2);
            Assert.Equal(2, Assert.Single(source.Redactions.Search(options).Areas).PageNumber);
            options.MatchCase = true;
            Assert.Empty(source.Redactions.Search(options).Areas);
            using CancellationTokenSource cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            options.CancellationToken = cancellation.Token;
            Assert.Throws<OperationCanceledException>(() => source.Redactions.Search(options));
        }

        [Fact]
        public void PreciseSelectionRejectsZeroLengthRegexAndLogicalKindCriteria() {
            PdfDocument source = PdfDocument.Load(BuildSingleTextObjectRedactionSource("(Alpha secret Omega) Tj"));
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Search(PreciseOptions().AddRegex("(?=secret)")));
            Assert.Throws<ArgumentException>(() => source.Redactions.Search(PreciseOptions().AddLogicalKind(PdfLogicalElementKind.TextBlock)));
        }

        [Theory]
        [InlineData("/Identity-V")]
        [InlineData("/UniJIS-UTF16-V")]
        public void PreciseApplyRefusesWholeObjectFallbackForUnsupportedGlyphRewrites(string encoding) {
            byte[] bytes = BuildType0ToUnicodeSingleTextObjectRedactionSource("Alpha secret Omega", encoding);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
            Assert.Equal(bytes, source.ToBytes());
        }

        [Fact]
        public void SearchKeepsExistingBlockDefaultAndAllowsExplicitTextOnlyPolicy() {
            PdfDocument source = PdfDocument.Load(BuildSingleTextObjectRedactionSource("(Alpha secret Omega) Tj"));
            PdfRedactionSearchOptions options = new PdfRedactionSearchOptions().AddLiteral("secret");
            PdfRedactionArea existing = Assert.Single(source.Redactions.Search(options).Areas);
            Assert.Equal(PdfRedactionContentScope.TextAndUnderlay, existing.ContentScope);
            options.ContentScope = PdfRedactionContentScope.TextOnly;
            PdfRedactionArea textOnly = Assert.Single(source.Redactions.Search(options).Areas);
            Assert.Equal(PdfRedactionContentScope.TextOnly, textOnly.ContentScope);
            PdfRedactionArea precise = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("secret")).Areas);
            Assert.True(precise.Width < existing.Width);
        }

        private static PdfRedactionSearchOptions PreciseOptions() => new() {
            TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
            ContentScope = PdfRedactionContentScope.TextOnly,
            RegexOptions = RegexOptions.CultureInvariant
        };

        private static string DescribeFindings(PdfRedactionPlan plan) => string.Join("; ", plan.Findings.Select(finding => finding.Code + ": " + finding.Message));

        private static IEnumerable<(string Text, double X, double Y, string Code, string Font)> ReadGlyphEvidence(byte[] bytes) =>
            PdfReadDocument.Open(bytes).Pages[0].GetGlyphTextSpans().SelectMany(span => {
                Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
                return glyphs.Select(glyph => (glyph.Text, Math.Round(glyph.X, 5), Math.Round(glyph.Y, 5),
                    Convert.ToBase64String(glyph.EncodedBytes.ToArray()), span.FontResource));
            });
    }
}
