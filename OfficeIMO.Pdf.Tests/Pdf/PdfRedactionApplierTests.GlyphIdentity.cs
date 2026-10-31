using OfficeIMO.Pdf;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfRedactionApplierTests {
        // 12 pt Helvetica at the 1.15 em leading of ordinary single-spaced text.
        private const string SingleSpaced = "BT /F1 12 Tf 13.8 TL 72 720 Td (The Revenue line above) Tj T* " +
            "(Quarterly invoice Revenue grew) Tj T* (Gypsy typography below) Tj ET";

        [Theory]
        [InlineData("invoice", 1)]
        [InlineData("Revenue", 2)]
        public void SearchRemovesOnlyMatchedGlyphsFromSingleSpacedText(string literal, int occurrences) {
            byte[] bytes = BuildTextContentRedactionSource(SingleSpaced);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral(literal));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            AssertRemovedExactly(bytes, result.Pdf, literal, occurrences);
            Assert.All(plan.Matches.Where(match => match.Kind == PdfRedactionMatchKind.TextBlock),
                match => Assert.Equal(literal, match.Text));
        }

        [Fact]
        public void SearchKeepsTheFullMarkOverDescendersAndAscenders() {
            PdfDocument source = PdfDocument.Load(BuildTextContentRedactionSource(SingleSpaced));
            PdfRedactionArea area = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("invoice")).Areas);
            Assert.Equal(15D, area.Height, 5);
            Assert.Equal(720D - 13.8D - 3D, area.Y, 5);
        }

        [Fact]
        public void RepeatedWordOnAdjacentLinesCanBeRemovedOneOccurrenceAtATime() {
            byte[] bytes = BuildTextContentRedactionSource(
                "BT /F1 12 Tf 13.8 TL 72 720 Td (secret first) Tj T* (secret second) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan search = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(search.IsReviewable, DescribeFindings(search));
            Assert.Equal(2, search.Areas.Count);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Plan(new[] { search.Areas[0] }));

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            string text = result.ToDocument().Read().Text;
            Assert.Single(Regex.Matches(text, "secret"));
            Assert.Contains("secret second", text, StringComparison.Ordinal);
            Assert.Contains("first", text, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("BT /F1 12 Tf 13.8 TL 72 720 Td (Alpha ) Tj (sec) Tj (ret) Tj ( Omega) Tj T* (next line) Tj ET")]
        [InlineData("BT /F1 12 Tf 13.8 TL 72 720 Td [(Alpha se) 40 (cr) -30 (et Omega)] TJ T* (next line) Tj ET")]
        [InlineData("BT /F1 12 Tf 13.8 TL 72 720 Td (Alpha ) ' (secret Omega) Tj 0 -13.8 Td (next line) Tj ET")]
        public void MatchesAcrossTextRunsAndKerningKeepTheirNeighbours(string content) {
            byte[] bytes = BuildTextContentRedactionSource(content);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            AssertRemovedExactly(bytes, result.Pdf, "secret");
            Assert.Contains("next line", result.ToDocument().Read().Text, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("BT /F1 12 Tf 14 TL 72 720 Td (Alpha secret) Tj T* (account Omega) Tj ET", "secret account")]
        [InlineData("BT /F1 20 Tf 72 720 Td [(Alpha ) 120 (secret) -120 ( Omega)] TJ ET", "secret")]
        public void SearchMarksMayTouchKeptTextWithoutRemovingIt(string content, string literal) {
            byte[] bytes = BuildTextContentRedactionSource(content);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral(literal));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            string text = result.ToDocument().Read().Text;
            Assert.Contains("Alpha", text, StringComparison.Ordinal);
            Assert.Contains("Omega", text, StringComparison.Ordinal);
        }

        [Fact]
        public void ManualRectangleOverTheSameGeometryStillRemovesTheNeighbouringLine() {
            byte[] bytes = BuildTextContentRedactionSource(SingleSpaced);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionArea mark = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("invoice")).Areas);
            var rectangle = new PdfRedactionArea(1, mark.X, mark.Y, mark.Width, mark.Height, "manual", PdfRedactionContentScope.TextOnly);

            string text = source.Redactions.Apply(new[] { rectangle }).Read().Text;

            Assert.DoesNotContain("invoice", text, StringComparison.Ordinal);
            Assert.DoesNotContain("line above", text.Replace("\n", " "), StringComparison.Ordinal);
        }

        [Fact]
        public void TightLeadingThatHidesTheNeighbourIsRefused() {
            byte[] bytes = BuildTextContentRedactionSource("BT /F1 12 Tf 8 TL 72 720 Td (Alpha above text) Tj T* (Alpha secret text) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-unselected-text-intersection");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
            Assert.Equal(bytes, source.ToBytes());
        }

        [Fact]
        public void DiagonalTextKeepsTheRefusalBecauseItsMarkWouldCoverTheNeighbour() {
            byte[] bytes = BuildTextContentRedactionSource(
                "BT /F1 12 Tf 0.7071 0.7071 -0.7071 0.7071 200 400 Tm 13.8 TL (A confidential word here) Tj T* (Another nearby line of text) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("confidential"));
            Assert.False(plan.IsReviewable, DescribeFindings(plan));
            Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-unselected-text-intersection");
        }

        [Theory]
        [InlineData("0 1 -1 0 200 300")]
        [InlineData("-1 0 0 -1 500 600")]
        public void RotatedSingleSpacedTextKeepsTheNeighbouringLine(string matrix) {
            byte[] bytes = BuildTextContentRedactionSource(
                "BT /F1 12 Tf " + matrix + " Tm 13.8 TL (Alpha secret Omega) Tj T* (Gypsy next line) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            AssertRemovedExactly(bytes, result.Pdf, "secret");
        }

        [Fact]
        public void IdenticalTextPaintedTwiceIsRemovedFromBothLayers() {
            byte[] bytes = BuildTextContentRedactionSource(
                "BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            if (!plan.IsReviewable) {
                Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-unselected-text-intersection");
                return;
            }
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.DoesNotContain("secret", result.ToDocument().Read().Text, StringComparison.Ordinal);
            Assert.DoesNotContain(ReadAllGlyphs(result.Pdf), glyph => glyph.Text == "s" && glyph.Y == 720D);
        }

        [Fact]
        public void BoldEffectFromTextPaintedTwiceWithAnOffsetIsRemovedFromBothCopies() {
            byte[] bytes = BuildTextContentRedactionSource(
                "BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET BT /F1 12 Tf 72.3 720 Td (Alpha secret Omega) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            AssertRemovedExactly(bytes, result.Pdf, "secret", expectedOccurrences: 2);
        }

        [Fact]
        public void MatchesInHiddenLayersAreRemovedWhereverTheyAre() {
            const string content = "BT /F1 12 Tf 13.8 TL 72 720 Td (Alpha secret Omega) Tj T* (Keep this line) Tj ET " +
                "/OC /Layer BDC BT /F1 12 Tf 72 400 Td (Layer secret note) Tj ET EMC";
            byte[] bytes = BuildLayeredSource(content);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Equal(2, plan.Areas.Count);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            AssertRemovedExactly(bytes, result.Pdf, "secret", expectedOccurrences: 2);
        }

        [Fact]
        public void FormXObjectInvokedTwiceLosesOnlyTheSelectedPlacement() {
            byte[] form = System.Text.Encoding.ASCII.GetBytes("BT /F1 12 Tf 0 0 Td (Form secret text) Tj ET");
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes("q 1 0 0 1 72 720 cm /Fm1 Do Q q 1 0 0 1 72 706.2 cm /Fm1 Do Q")),
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 -20 300 20] /Resources << /Font << /F1 4 0 R >> >> /Length " + form.Length + " >>\nstream\n" +
                    System.Text.Encoding.ASCII.GetString(form) + "\nendstream\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan search = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(search.IsReviewable, DescribeFindings(search));
            Assert.Equal(2, search.Areas.Count);

            PdfRedactionApplyResult one = source.Redactions.ApplyWithEvidence(source.Redactions.Plan(new[] { search.Areas[0] }));
            Assert.True(one.Evidence.IsVerified, one.Evidence.Summary);
            Assert.Single(Regex.Matches(one.ToDocument().Read().Text, "secret"));
            Assert.Equal(2, Regex.Matches(one.ToDocument().Read().Text, "Form").Count);

            PdfRedactionApplyResult both = source.Redactions.ApplyWithEvidence(search);
            Assert.True(both.Evidence.IsVerified, both.Evidence.Summary);
            AssertRemovedExactly(bytes, both.Pdf, "secret", expectedOccurrences: 2);
            Assert.DoesNotContain("Form secret", PdfEncoding.Latin1GetString(both.Pdf), StringComparison.Ordinal);
            Assert.Empty(PdfReadDocument.Open(both.Pdf).Objects.Values.Select(item => item.Value).OfType<PdfStream>()
                .Where(stream => PdfEncoding.Latin1GetString(OfficeIMO.Pdf.Filters.StreamDecoder.DecodeRequired(stream, PdfReadDocument.Open(both.Pdf).Objects, int.MaxValue)).IndexOf("secret", StringComparison.Ordinal) >= 0));
        }

        [Fact]
        public void TaggedAlternateTextHoldingTheMatchIsRemovedAndOtherDescriptionsStay() {
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET")),
                "6 0 obj\n<< /Type /StructTreeRoot /K [7 0 R 8 0 R] >>\nendobj",
                "7 0 obj\n<< /Type /StructElem /S /P /P 6 0 R /Alt (the secret alt) /ActualText (Alpha secret actual) >>\nendobj",
                "8 0 obj\n<< /Type /StructElem /S /P /P 6 0 R /Alt (A harmless picture) >>\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionArea mark = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("secret")).Areas);

            string raw = PdfEncoding.Latin1GetString(source.Redactions.Apply(source.Redactions.Plan(new[] { mark })).ToBytes());

            Assert.DoesNotContain("secret", raw, StringComparison.Ordinal);
            Assert.Contains("A harmless picture", raw, StringComparison.Ordinal);
        }

        [Fact]
        public void SelectionMustBelongToTheDocumentItIsAppliedTo() {
            PdfDocument first = PdfDocument.Load(BuildTextContentRedactionSource(SingleSpaced));
            PdfRedactionArea mark = Assert.Single(first.Redactions.Search(PreciseOptions().AddLiteral("invoice")).Areas);
            byte[] otherBytes = BuildTextContentRedactionSource(SingleSpaced.Replace("invoice", "pricing"));
            PdfDocument other = PdfDocument.Load(otherBytes);

            PdfRedactionPlan plan = other.Redactions.Plan(new[] { mark });

            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-selection-stale");
            Assert.Throws<InvalidOperationException>(() => other.Redactions.Apply(new[] { mark }));
            Assert.Equal(otherBytes, other.ToBytes());
        }

        [Fact]
        public void DirectAreaRouteAppliesSearchMarksByIdentity() {
            byte[] bytes = BuildTextContentRedactionSource(SingleSpaced);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionArea mark = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("invoice")).Areas);
            AssertRemovedExactly(bytes, source.Redactions.Apply(new[] { mark }).ToBytes(), "invoice");
        }

        [Fact]
        public void SearchMarksCannotBecomeRectangleAnnotations() {
            PdfDocument source = PdfDocument.Load(BuildTextContentRedactionSource(SingleSpaced));
            PdfRedactionArea mark = Assert.Single(source.Redactions.Search(PreciseOptions().AddLiteral("invoice")).Areas);
            Assert.Throws<InvalidOperationException>(() => source.Redactions.AddAnnotation(
                new PdfRedactionAnnotationOptions(PdfRedactionRegion.Group(1, new[] { mark }))));
        }

        [Fact]
        public void ImageUnderAMatchIsWipedAcrossTheFullMarkAndNeighbouringTextStays() {
            byte[] bytes = BuildScanWithTextLayer(out string content);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionSearchOptions options = PreciseOptions().AddLiteral("secret");
            options.ContentScope = PdfRedactionContentScope.TextAndUnderlay;
            PdfRedactionPlan plan = source.Redactions.Search(options);
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.ImagePlacement);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            AssertRemovedExactly(bytes, result.Pdf, "secret");
            Assert.Contains(result.Evidence.Items, item => item.ReviewedMatch.Kind == PdfRedactionMatchKind.ImagePlacement);
        }

        [Fact]
        public void KeepingTheImageUnderAMatchWithNearbyTextIsRefused() {
            byte[] bytes = BuildScanWithTextLayer(out _);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-unselected-text-intersection");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
            Assert.Equal(bytes, source.ToBytes());
        }

        [Theory]
        [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
        [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
        public void EncryptedSingleSpacedSourceStaysEncrypted(PdfStandardEncryptionAlgorithm algorithm) {
            byte[] bytes = BuildTextContentRedactionSource(SingleSpaced);
            byte[] encrypted = PdfDocument.Load(bytes).Security.Encrypt(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner", Algorithm = algorithm }).Pdf;
            PdfDocument source = PdfDocument.Load(encrypted, new PdfLoadOptions { Password = "owner" });
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("invoice"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));

            byte[] output = source.Redactions.Apply(plan).ToBytes();

            Assert.ThrowsAny<Exception>(() => PdfReadDocument.Open(output));
            PdfDocument opened = PdfDocument.Load(output, new PdfLoadOptions { Password = "owner" });
            Assert.DoesNotContain("invoice", opened.Read().Text, StringComparison.Ordinal);
            Assert.Contains("line above", opened.Read().Text.Replace("\n", " "), StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        public void TouchingMarksOnAdjacentLinesDoNotCountAsResidualContent(PdfRedactionTextSelection selection) {
            byte[] bytes = BuildTextContentRedactionSource(
                "BT /F1 12 Tf 13.8 TL 72 720 Td (secret first) Tj T* (secret second) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Equal(2, plan.Areas.Count);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.DoesNotContain("secret", result.ToDocument().Read().Text, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("(hidden secret words)")]
        [InlineData("<FEFF00730065006300720065007400>")]
        public void WholeBlockSearchClearsMarkedContentActualText(string value) {
            byte[] bytes = BuildTextContentRedactionSource("/Span << /ActualText " + value + " >> BDC BT /F1 12 Tf 72 720 Td (X) Tj ET EMC " +
                "BT /F1 12 Tf 72 500 Td (Kept text) Tj ET");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));

            byte[] output = source.Redactions.Apply(plan).ToBytes();

            string raw = DecodeAllStreams(output);
            Assert.DoesNotContain("ActualText", raw, StringComparison.Ordinal);
            Assert.Contains("Kept text", raw, StringComparison.Ordinal);
        }

        [Fact]
        public void WholeBlockSearchKeepsTheNeighbouringLine() {
            byte[] bytes = BuildTextContentRedactionSource(SingleSpaced);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("invoice"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            string text = string.Concat(ReadAllGlyphs(result.Pdf).Select(static glyph => glyph.Text));
            Assert.DoesNotContain("Quarterly", text, StringComparison.Ordinal);
            Assert.Contains("The Revenue line above", text, StringComparison.Ordinal);
            Assert.Contains("Gypsy typography below", text, StringComparison.Ordinal);
        }

        [Fact]
        public void WholeBlockAreaStaysABlockThroughARebuiltPlan() {
            byte[] bytes = BuildTextContentRedactionSource(SingleSpaced);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionArea mark = Assert.Single(source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("invoice")).Areas);
            string text = string.Concat(ReadAllGlyphs(source.Redactions.Apply(new[] { mark.WithLabel("Reviewed") }).ToBytes()).Select(static glyph => glyph.Text));
            Assert.Contains("Gypsy typography below", text, StringComparison.Ordinal);
            Assert.DoesNotContain("grew", text, StringComparison.Ordinal);
        }

        [Fact]
        public void WholeBlockSearchRefusesTextOverlappingTheBlockThatIsNotPartOfIt() {
            byte[] bytes = BuildLayeredSource("BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET /OC /Layer BDC BT /F1 12 Tf 72 720 Td (private) Tj ET EMC");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("secret"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-unselected-text-intersection");
            Assert.Contains(plan.Findings, finding => finding.Message.Contains("would remove an unselected glyph"));
            Assert.Throws<InvalidOperationException>(() => source.Redactions.Apply(plan));
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        public void HiddenMatchesKeepOverlappingVisibleContentAndPaintNoMark(PdfRedactionTextSelection selection) {
            byte[] bytes = BuildLayeredSource("q 0.8 g 60 290 300 30 re f Q BT /F1 12 Tf 72 300 Td (Visible words stay here) Tj ET " +
                "/OC /Layer BDC BT /F1 12 Tf 72 300 Td (Hidden secret note) Tj ET EMC");
            PdfDocument source = PdfDocument.Load(bytes);
            var options = new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret");
            PdfRedactionPlan plan = source.Redactions.Search(options);
            Assert.True(plan.IsReviewable, DescribeFindings(plan));

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            string text = string.Concat(ReadAllGlyphs(result.Pdf).Select(static glyph => glyph.Text));
            Assert.Contains("Visible words stay here", text, StringComparison.Ordinal);
            Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
            string raw = DecodeAllStreams(result.Pdf);
            Assert.Contains("60 290 300 30 re f", raw, StringComparison.Ordinal);
            Assert.Single(PdfReadDocument.Open(result.Pdf).Pages[0].GetIdentityVisualPrimitives());
        }

        [Fact]
        public void HiddenMatchSharingBytesAndPositionWithVisibleTextIsRefused() {
            byte[] bytes = BuildLayeredSource("BT /F1 12 Tf 72 300 Td (sacred) Tj ET /OC /Layer BDC BT /F1 12 Tf 72 300 Td (secret) Tj ET EMC");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, finding => finding.Code == "redaction-search-unselected-text-intersection");
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void WholeBlockSearchRemovesHiddenLayerMatches(bool regex) {
            byte[] bytes = BuildLayeredSource("BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET BT /F1 12 Tf 72 500 Td (Unrelated visible text) Tj ET " +
                "/OC /Layer BDC BT /F1 12 Tf 72 300 Td (Layer secret note) Tj ET BT /F1 12 Tf 72 200 Td (Other hidden words) Tj ET EMC");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionSearchOptions options = new PdfRedactionSearchOptions();
            if (regex) options.AddRegex("sec[r]et"); else options.AddLiteral("secret");
            PdfRedactionPlan plan = source.Redactions.Search(options);
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Equal(2, plan.Areas.Count);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            IReadOnlyList<(string Text, double X, double Y)> glyphs = ReadAllGlyphs(result.Pdf);
            string kept = string.Concat(glyphs.Select(static glyph => glyph.Text));
            Assert.DoesNotContain("secret", kept, StringComparison.Ordinal);
            Assert.DoesNotContain("Layer", kept, StringComparison.Ordinal);
            Assert.Contains("Unrelated visible text", kept, StringComparison.Ordinal);
            Assert.Contains("Other hidden words", kept, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks, false)]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks, true)]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs, false)]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs, true)]
        public void AlternateTextHoldingTheSearchIsRemovedThroughARebuiltPlan(PdfRedactionTextSelection selection, bool regex) {
            byte[] bytes = BuildTaggedSource();
            PdfDocument source = PdfDocument.Load(bytes);
            var options = new PdfRedactionSearchOptions { TextSelection = selection };
            if (regex) options.AddRegex("sec[r]et"); else options.AddLiteral("secret");
            PdfRedactionArea mark = Assert.Single(source.Redactions.Search(options).Areas);

            string raw = PdfEncoding.Latin1GetString(source.Redactions.Apply(source.Redactions.Plan(new[] { mark.WithLabel("Reviewed") })).ToBytes());

            Assert.DoesNotContain("secret", raw, StringComparison.Ordinal);
            Assert.Contains("A harmless picture", raw, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        public void WholeBlockAndGlyphSearchClearFormsPaintedTwice(PdfRedactionTextSelection selection) {
            byte[] bytes = BuildRepeatedFormSource();
            PdfDocument source = PdfDocument.Load(bytes);
            var options = new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret");
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Search(options));
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.DoesNotContain("Form secret", DecodeAllStreams(result.Pdf), StringComparison.Ordinal);
        }

        [Fact]
        public void VerifierFlagsAnOrphanedFormHoldingTheSearchedText() {
            byte[] bytes = BuildRepeatedFormSource();
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            // Same document, but the page no longer paints the form that holds the text.
            byte[] orphaned = BuildRepeatedFormSource("q 1 0 0 1 72 720 cm Q");
            Assert.DoesNotContain("secret", PdfDocument.Load(orphaned).Read().Text, StringComparison.Ordinal);

            PdfRedactionVerificationReport report = PdfDocument.Load(orphaned).Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());

            Assert.Contains(report.Issues, issue => issue.Feature == "RedactedContentOutsidePageContent" && issue.Marker == "OrphanedFormXObject");
        }

        [Fact]
        public void VerifierAcceptsAFormStillPaintedAndFlagsRetainedAlternateText() {
            byte[] bytes = BuildRepeatedFormSource();
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan search = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            PdfRedactionApplyResult one = source.Redactions.ApplyWithEvidence(source.Redactions.Plan(new[] { search.Areas[0] }));
            Assert.True(one.Evidence.IsVerified, one.Evidence.Summary);

            PdfDocument tagged = PdfDocument.Load(BuildTaggedSource());
            PdfRedactionPlan plan = tagged.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            PdfRedactionVerificationReport report = tagged.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.Contains(report.Issues, issue => issue.Feature == "RedactedContentOutsidePageContent" && issue.Marker == "AlternateText");
        }

        // Two placements of one form, each drawing the same marked image.
        private byte[] BuildRepeatedImageFormSource(string? pageContent = null) {
            const int width = 40, height = 40;
            byte[] pixels = Enumerable.Repeat((byte)0x5A, width * height).ToArray();
            const string form = "q 100 0 0 100 0 0 cm /Im1 Do Q";
            return BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(pageContent ?? "q 1 0 0 1 100 600 cm /Fm1 Do Q q 1 0 0 1 100 300 cm /Fm1 Do Q")),
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Im1 7 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj",
                "7 0 obj\n<< /Type /XObject /Subtype /Image /Width " + width + " /Height " + height + " /ColorSpace /DeviceGray /BitsPerComponent 8 /Length " + pixels.Length +
                    " >>\nstream\n" + PdfEncoding.Latin1GetString(pixels) + "\nendstream\nendobj"
            }, rootObjectNumber: 1);
        }

        private static int CountUntouchedImagePixels(byte[] pdf) {
            Dictionary<int, PdfIndirectObject> objects = PdfReadDocument.Open(pdf).Objects;
            return objects.Values.Select(static item => item.Value).OfType<PdfStream>()
                .Where(static stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image")
                .Select(stream => OfficeIMO.Pdf.Filters.StreamDecoder.DecodeRequired(stream, objects, int.MaxValue))
                .Count(static data => data.All(static value => value == 0x5A));
        }

        [Fact]
        public void RedactingEveryPlacementOfARepeatedImageFormLeavesNoOriginalPixels() {
            byte[] bytes = BuildRepeatedImageFormSource();
            Assert.Equal(1, CountUntouchedImagePixels(bytes));
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] {
                new PdfRedactionArea(1, 120, 620, 40, 40), new PdfRedactionArea(1, 120, 320, 40, 40)
            });

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Equal(0, CountUntouchedImagePixels(result.Pdf));
        }

        [Fact]
        public void RedactingOnePlacementKeepsTheOtherPaintedFromTheOriginalImage() {
            byte[] bytes = BuildRepeatedImageFormSource();
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) }));
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Equal(1, CountUntouchedImagePixels(result.Pdf));
            Assert.Equal(2, PdfReadDocument.Open(result.Pdf).Pages[0].GetImagePlacements().Count);
        }

        [Fact]
        public void VerifierFlagsAFormNothingPaintsThatHoldsTheRedactedImage() {
            PdfDocument source = PdfDocument.Load(BuildRepeatedImageFormSource());
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) });
            PdfDocument orphaned = PdfDocument.Load(BuildRepeatedImageFormSource("q Q"));

            PdfRedactionVerificationReport report = orphaned.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());

            Assert.Contains(report.Issues, issue => issue.Feature == "RedactedContentOutsidePageContent" && issue.Marker == "OrphanedImageData");
        }

        [Fact]
        public void TextOnlyPlanDoesNotTreatItsKeptImageAsALeftover() {
            PdfDocument source = PdfDocument.Load(BuildRepeatedImageFormSource());
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40, null, PdfRedactionContentScope.TextOnly) });
            PdfDocument unpainted = PdfDocument.Load(BuildRepeatedImageFormSource("q Q"));
            PdfRedactionVerificationReport report = unpainted.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.DoesNotContain(report.Issues, issue => issue.Marker == "OrphanedImageData");
        }

        [Fact]
        public void RedactingAnImagePaintedTwiceDirectlyLeavesNoOriginalPixels() {
            byte[] bytes = BuildRepeatedImageFormSource("q 100 0 0 100 100 600 cm /Im1 Do Q q 100 0 0 100 100 300 cm /Im1 Do Q")
                ;
            bytes = PdfEncoding.Latin1GetBytes(PdfEncoding.Latin1GetString(bytes).Replace("/XObject << /Fm1 6 0 R >> >> >>", "/XObject << /Im1 7 0 R >> >> >>"));
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40), new PdfRedactionArea(1, 120, 320, 40, 40) });
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Equal(0, CountUntouchedImagePixels(result.Pdf));
        }

        [Fact]
        public void RedactingANestedRepeatedImageFormLeavesNoOriginalPixels() {
            const int size = 40;
            byte[] pixels = Enumerable.Repeat((byte)0x5A, size * size).ToArray();
            const string inner = "q 100 0 0 100 0 0 cm /Im1 Do Q", outer = "q 1 0 0 1 0 0 cm /In1 Do Q q 1 0 0 1 200 0 cm /In1 Do Q";
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /XObject << /Fm1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes("q 1 0 0 1 100 600 cm /Fm1 Do Q")),
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 400 100] /Resources << /XObject << /In1 8 0 R >> >> /Length " + outer.Length + " >>\nstream\n" + outer + "\nendstream\nendobj",
                "7 0 obj\n<< /Type /XObject /Subtype /Image /Width " + size + " /Height " + size + " /ColorSpace /DeviceGray /BitsPerComponent 8 /Length " + pixels.Length + " >>\nstream\n" + PdfEncoding.Latin1GetString(pixels) + "\nendstream\nendobj",
                "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Im1 7 0 R >> >> /Length " + inner.Length + " >>\nstream\n" + inner + "\nendstream\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40), new PdfRedactionArea(1, 320, 620, 40, 40) });
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Equal(0, CountUntouchedImagePixels(result.Pdf));
        }

        [Fact]
        public void RedactingAFormSharedByTwoPagesLeavesNoOriginalPixels() {
            const int size = 40;
            byte[] pixels = Enumerable.Repeat((byte)0x5A, size * size).ToArray();
            const string form = "q 100 0 0 100 0 0 cm /Im1 Do Q";
            byte[] content = System.Text.Encoding.ASCII.GetBytes("q 1 0 0 1 100 600 cm /Fm1 Do Q");
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /XObject << /Fm1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /XObject << /Fm1 6 0 R >> >> >>\nendobj",
                BuildStreamObject(5, content),
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Im1 7 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj",
                "7 0 obj\n<< /Type /XObject /Subtype /Image /Width " + size + " /Height " + size + " /ColorSpace /DeviceGray /BitsPerComponent 8 /Length " + pixels.Length + " >>\nstream\n" + PdfEncoding.Latin1GetString(pixels) + "\nendstream\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40), new PdfRedactionArea(2, 120, 620, 40, 40) });
            PdfRedactionApplyResult both = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(both.Evidence.IsVerified, both.Evidence.Summary);
            Assert.Equal(0, CountUntouchedImagePixels(both.Pdf));

            PdfRedactionApplyResult one = source.Redactions.ApplyWithEvidence(source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) }));
            Assert.True(one.Evidence.IsVerified, one.Evidence.Summary);
            Assert.Equal(1, CountUntouchedImagePixels(one.Pdf));
        }

        private static string GrayImage(int number, string extra = "", byte value = 0x5A) {
            const int size = 40;
            string pixels = new string((char)value, size * size);
            return number + " 0 obj\n<< /Type /XObject /Subtype /Image /Width " + size + " /Height " + size + " /ColorSpace /DeviceGray /BitsPerComponent 8 " + extra +
                " /Length " + pixels.Length + " >>\nstream\n" + pixels + "\nendstream\nendobj";
        }

        private byte[] BuildImagePage(string pageResources, string pageContent, params string[] extra) => BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << " + pageResources + " >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(pageContent))
        }.Concat(extra).ToArray(), rootObjectNumber: 1);

        [Theory]
        [InlineData("/SMask 7 0 R")]
        [InlineData("/Mask 7 0 R")]
        public void MaskDataOfARedactedImageDoesNotSurviveUntouched(string link) {
            byte[] bytes = BuildImagePage("/XObject << /Im1 6 0 R >>", "q 100 0 0 100 100 600 cm /Im1 Do Q",
                GrayImage(6, link), GrayImage(7, link.StartsWith("/Mask", StringComparison.Ordinal) ? "/ImageMask true" : string.Empty, 0x77));
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) });
            try {
                PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
                Dictionary<int, PdfIndirectObject> objects = PdfReadDocument.Open(result.Pdf).Objects;
                Assert.DoesNotContain(objects.Values.Select(static item => item.Value).OfType<PdfStream>(), stream =>
                    OfficeIMO.Pdf.Filters.StreamDecoder.DecodeRequired(stream, objects, int.MaxValue) is { Length: > 100 } data && data.All(static value => value == 0x77));
            } catch (InvalidOperationException) {
                // Refusing is also safe.
            }
        }

        [Fact]
        public void InlineImageUnderAMarkIsRefusedNotKept() {
            string pixels = new string('Z', 400);
            byte[] bytes = BuildImagePage("", "q 100 0 0 100 100 600 cm BI /W 20 /H 20 /CS /G /BPC 8 ID " + pixels + " EI Q");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) });
            Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.ImagePlacement);
            Exception? refused = Record.Exception(() => source.Redactions.ApplyWithEvidence(plan));
            if (refused is null) Assert.DoesNotContain(pixels, DecodeAllStreams(source.Redactions.ApplyWithEvidence(plan).Pdf), StringComparison.Ordinal);
            else Assert.IsType<InvalidOperationException>(refused);
        }

        [Fact]
        public void VerifierSeesAnUnpaintedFormHoldingAnInlineImageWhenOneWasReviewed() {
            string pixels = new string('Z', 400);
            string inline = "BI /W 20 /H 20 /CS /G /BPC 8 ID " + pixels + " EI";
            string form = "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Length " + inline.Length + " >>\nstream\n" + inline + "\nendstream\nendobj";
            PdfDocument source = PdfDocument.Load(BuildImagePage("", "q 100 0 0 100 100 600 cm " + inline + " Q"));
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) });
            PdfDocument output = PdfDocument.Load(BuildImagePage("/XObject << /Fm1 6 0 R >>", "q Q", form));
            PdfRedactionVerificationReport report = output.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.Contains(report.Issues, issue => issue.Marker == "OrphanedImageData");
        }

        [Fact]
        public void VerifierSeesAMaskOfAReviewedImageHeldByAnUnpaintedResource() {
            PdfDocument source = PdfDocument.Load(BuildImagePage("/XObject << /Im1 6 0 R >>", "q 100 0 0 100 100 600 cm /Im1 Do Q", GrayImage(6, "/SMask 7 0 R"), GrayImage(7, value: 0x77)));
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 120, 620, 40, 40) });
            PdfDocument output = PdfDocument.Load(BuildImagePage("/XObject << /Other 6 0 R >>", "q Q", GrayImage(6, "/SMask 7 0 R", 0x10), GrayImage(7, value: 0x77)));
            PdfRedactionVerificationReport report = output.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.Contains(report.Issues, issue => issue.Marker == "OrphanedImageData");
        }

        [Theory]
        [InlineData("pattern")]
        [InlineData("softmask")]
        public void VerifierSeesReviewedImageHeldByPatternsAndSoftMaskGroups(string holder) {
            const string form = "q 100 0 0 100 0 0 cm /Im1 Do Q";
            string extraResources = holder == "pattern" ? "/Pattern << /P1 8 0 R >>" : "/ExtGState << /G1 << /SMask << /S /Luminosity /G 8 0 R >> >> >>";
            string holderObject = holder == "pattern"
                ? "8 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /XObject << /Im1 6 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj"
                : "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /S /Transparency >> /Resources << /XObject << /Im1 6 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj";
            byte[] painted = BuildImagePage("/XObject << /Im1 6 0 R >> " + extraResources, "q 100 0 0 100 100 600 cm /Im1 Do Q", GrayImage(6), holderObject);
            PdfDocument source = PdfDocument.Load(painted);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 100, 600, 100, 100) });
            byte[] output = BuildImagePage(extraResources, "q Q", GrayImage(6), holderObject);

            PdfRedactionVerificationReport report = PdfDocument.Load(output).Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());

            Assert.Contains(report.Issues, issue => issue.Marker == "OrphanedImageData");
        }

        [Fact]
        public void VerifierSeesReviewedImageHeldByAnAnnotationAppearance() {
            const string form = "q 100 0 0 100 0 0 cm /Im1 Do Q";
            string appearance = "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Im1 6 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj";
            string stamp = "9 0 obj\n<< /Type /Annot /Subtype /Stamp /Rect [300 100 400 200] /AP << /N 8 0 R >> >>\nendobj";
            byte[] painted = BuildImagePage("/XObject << /Im1 6 0 R >>", "q 100 0 0 100 100 600 cm /Im1 Do Q", GrayImage(6), appearance, stamp);
            PdfDocument source = PdfDocument.Load(painted);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 100, 600, 100, 100) });
            byte[] output = PdfEncoding.Latin1GetBytes(PdfEncoding.Latin1GetString(BuildImagePage("", "q Q", GrayImage(6), appearance, stamp))
                .Replace("/Contents 5 0 R", "/Contents 5 0 R /Annots [9 0 R]"));
            PdfRedactionVerificationReport report = PdfDocument.Load(output).Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.Contains(report.Issues, issue => issue.Marker == "OrphanedImageData");
        }

        private byte[] BuildAnnotatedPage(string content, string annotationObjects, string annots, params string[] extra) => BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Annots [" + annots + "] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(content)),
            annotationObjects
        }.Concat(extra).ToArray(), rootObjectNumber: 1);

        // Line boxes at 720 and 706.2: a rule in the gap and a link on the next line touch the mark, not the glyphs.
        private const string TwoLines = "BT /F1 12 Tf 13.8 TL 72 720 Td (Alpha secret Omega) Tj T* (Linked next line) Tj ET";

        [Theory]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        public void NeighbouringLinkAndTableRuleStayUnderTheMark(PdfRedactionTextSelection selection) {
            byte[] bytes = BuildAnnotatedPage("60 716.9 300 0.4 re f " + TwoLines,
                "6 0 obj\n<< /Type /Annot /Subtype /Link /Rect [72 703 200 717.2] /Border [0 0 0] /A << /S /URI /URI (https://example.com/) >> >>\nendobj", "6 0 R");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.DoesNotContain(plan.Matches, match => match.Kind == PdfRedactionMatchKind.VectorPath || match.Kind == PdfRedactionMatchKind.Annotation);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Contains("60 716.9 300 0.4 re f", DecodeAllStreams(result.Pdf), StringComparison.Ordinal);
            Assert.Contains("https://example.com/", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
        }

        [Fact]
        public void PathThroughTheMatchedGlyphsAndAnnotationOverThemStillGo() {
            byte[] bytes = BuildAnnotatedPage("70 722 200 2 re f " + TwoLines,
                "6 0 obj\n<< /Type /Annot /Subtype /Highlight /Rect [100 718 140 730] >>\nendobj", "6 0 R");
            PdfDocument source = PdfDocument.Load(bytes);
            var options = new PdfRedactionSearchOptions { TextSelection = PdfRedactionTextSelection.MatchedGlyphs }.AddLiteral("secret");
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Search(options));
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.DoesNotContain("70 722 200 2 re f", DecodeAllStreams(result.Pdf), StringComparison.Ordinal);
            Assert.DoesNotContain("/Highlight", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        public void AnnotationCarryingTheTextGoesWhereverItSitsAndOthersStay(PdfRedactionTextSelection selection) {
            const string appearance = "BT /F1 12 Tf 0 0 Td (Stamp secret words) Tj ET";
            byte[] bytes = BuildAnnotatedPage(TwoLines,
                "6 0 obj\n<< /Type /Annot /Subtype /Text /Rect [400 100 420 120] /Contents (About the secret plan) >>\nendobj",
                "6 0 R 7 0 R 9 0 R",
                "7 0 obj\n<< /Type /Annot /Subtype /Stamp /Rect [400 200 500 220] /AP << /N 8 0 R >> >>\nendobj",
                "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Resources << /Font << /F1 4 0 R >> >> /Length " + appearance.Length + " >>\nstream\n" + appearance + "\nendstream\nendobj",
                "9 0 obj\n<< /Type /Annot /Subtype /Text /Rect [400 300 420 320] /Contents (Harmless remark) >>\nendobj");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            string raw = PdfEncoding.Latin1GetString(result.Pdf) + DecodeAllStreams(result.Pdf);
            Assert.DoesNotContain("secret", raw, StringComparison.Ordinal);
            Assert.Contains("Harmless remark", raw, StringComparison.Ordinal);
        }

        [Fact]
        public void VerifierFlagsAnAnnotationStillCarryingTheText() {
            PdfDocument source = PdfDocument.Load(BuildTextContentRedactionSource(TwoLines));
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            PdfDocument output = PdfDocument.Load(BuildAnnotatedPage("q Q", "6 0 obj\n<< /Type /Annot /Subtype /Text /Rect [400 100 420 120] /Contents (The secret plan) >>\nendobj", "6 0 R"));
            PdfRedactionVerificationReport report = output.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.Contains(report.Issues, issue => issue.Marker == "AnnotationText");
        }

        [Fact]
        public void ManualRectangleStillRemovesTheAnnotationAndPathItTouches() {
            byte[] bytes = BuildAnnotatedPage("60 716.9 300 0.4 re f " + TwoLines,
                "6 0 obj\n<< /Type /Annot /Subtype /Link /Rect [72 703 200 717.2] /A << /S /URI /URI (https://example.com/) >> >>\nendobj", "6 0 R");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Plan(new[] { new PdfRedactionArea(1, 100, 715, 100, 15) }));
            Assert.DoesNotContain("https://example.com/", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
            Assert.DoesNotContain("60 716.9 300 0.4 re f", DecodeAllStreams(result.Pdf), StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        public void UnpaintedFormInTheSourceHoldingTheTextIsCleaned(PdfRedactionTextSelection selection) {
            const string form = "BT /F1 12 Tf 0 0 Td (Form secret text) Tj ET";
            byte[] bytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Unused 6 0 R /Other 7 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(TwoLines)),
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 300 20] /Resources << /Font << /F1 4 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj",
                "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 300 20] /Length 3 >>\nstream\nq Q\nendstream\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.DoesNotContain("Form secret", DecodeAllStreams(result.Pdf), StringComparison.Ordinal);
            Assert.Contains("/Other", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
        }

        // Page 1 has the text. Page 2 has a note, page 3 an appearance, page 4 a widget and a harmless note.
        private byte[] BuildAnnotatedPages() {
            const string appearance = "BT /F1 12 Tf 0 0 Td (Stamp secret words) Tj ET";
            const string blank = "BT /F1 12 Tf 72 720 Td (Plain page) Tj ET";
            string Page(int number, int contents, string annots) => number + " 0 obj\n<< /Type /Page /Parent 2 0 R /Contents " + contents + " 0 R /Annots [" + annots + "] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj";
            return BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 4 /Kids [3 0 R 6 0 R 7 0 R 8 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                Page(3, 5, string.Empty),
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(TwoLines)),
                Page(6, 9, "10 0 R"), Page(7, 9, "11 0 R"), Page(8, 9, "13 0 R 14 0 R"),
                BuildStreamObject(9, System.Text.Encoding.ASCII.GetBytes(blank)),
                "10 0 obj\n<< /Type /Annot /Subtype /Text /Rect [400 100 420 120] /Contents (About the secret plan) >>\nendobj",
                "11 0 obj\n<< /Type /Annot /Subtype /Stamp /Rect [400 200 500 220] /AP << /N 12 0 R >> >>\nendobj",
                "12 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Resources << /Font << /F1 4 0 R >> >> /Length " + appearance.Length + " >>\nstream\n" + appearance + "\nendstream\nendobj",
                "13 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Tx /T (Field) /V (secret value) /Rect [100 100 200 120] >>\nendobj",
                "14 0 obj\n<< /Type /Annot /Subtype /Text /Rect [400 300 420 320] /Contents (Harmless remark) >>\nendobj"
            }, rootObjectNumber: 1);
        }

        [Theory]
        [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
        [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
        public void AnnotationsOnPagesWithoutATextMatchAreReportedAndRemoved(PdfRedactionTextSelection selection) {
            byte[] bytes = BuildAnnotatedPages();
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));
            Assert.Equal(new[] { 2, 3 }, plan.Matches.Where(static match => match.Kind == PdfRedactionMatchKind.Annotation).Select(static match => match.PageNumber).OrderBy(static page => page));
            Assert.DoesNotContain(plan.Areas, area => area.PageNumber == 4);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            string raw = PdfEncoding.Latin1GetString(result.Pdf);
            Assert.DoesNotContain("secret plan", raw, StringComparison.Ordinal);
            Assert.DoesNotContain("Stamp secret", raw + DecodeAllStreams(result.Pdf), StringComparison.Ordinal);
            Assert.Contains("Harmless remark", raw, StringComparison.Ordinal);
            Assert.Contains("secret value", raw, StringComparison.Ordinal);
            Assert.Equal(3, Regex.Matches(result.ToDocument().Read().Text, "Plain page").Count);
            Assert.Single(PdfReadDocument.Open(result.Pdf).Pages[0].GetIdentityVisualPrimitives());
            Assert.Empty(PdfReadDocument.Open(result.Pdf).Pages[1].GetIdentityVisualPrimitives());
        }

        [Fact]
        public void PageLimitedSearchLeavesAnnotationsOnOtherPages() {
            PdfDocument source = PdfDocument.Load(BuildAnnotatedPages());
            PdfRedactionSearchOptions options = PreciseOptions().AddLiteral("secret");
            options.PageNumbers.Add(1);
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Search(options));
            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Contains("secret plan", PdfEncoding.Latin1GetString(result.Pdf), StringComparison.Ordinal);
        }

        [Fact]
        public void VerifierFlagsAnAnnotationLeftOnAPageWithoutATextMatch() {
            byte[] bytes = BuildAnnotatedPages();
            PdfRedactionPlan plan = PdfDocument.Load(bytes).Redactions.Search(PreciseOptions().AddLiteral("secret"));
            PdfRedactionVerificationReport report = PdfDocument.Load(bytes).Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions());
            Assert.Contains(report.Issues, issue => issue.Marker == "AnnotationText" );
        }

        [Fact]
        public void MappingRefusalNoLongerPointsAtBlockSelection() {
            byte[] bytes = BuildTextContentRedactionSource("BT /F1 12 Tf 13.8 TL 72 720 Td /Span << /ActualText (hidden secret words) >> BDC (X) Tj EMC T* (Kept next line) Tj ET");
            PdfRedactionPlan plan = PdfDocument.Load(bytes).Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("secret"));
            PdfDiagnosticFinding finding = Assert.Single(plan.Findings, static item => item.Code == "redaction-search-glyph-mapping-unsupported");
            Assert.DoesNotContain("block", finding.Message, StringComparison.OrdinalIgnoreCase);
        }

        private byte[] BuildTaggedSource() => BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Alpha secret Omega) Tj ET")),
            "6 0 obj\n<< /Type /StructTreeRoot /K [7 0 R 8 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /StructElem /S /P /P 6 0 R /Alt (the secret alt) /ActualText (Alpha secret actual) >>\nendobj",
            "8 0 obj\n<< /Type /StructElem /S /P /P 6 0 R /Alt (A harmless picture) >>\nendobj"
        }, rootObjectNumber: 1);

        private byte[] BuildRepeatedFormSource(string? pageContent = null) {
            const string form = "BT /F1 12 Tf 0 0 Td (Form secret text) Tj ET";
            return BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(pageContent ?? "q 1 0 0 1 72 720 cm /Fm1 Do Q q 1 0 0 1 72 680 cm /Fm1 Do Q")),
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 -20 300 20] /Resources << /Font << /F1 4 0 R >> >> /Length " + form.Length + " >>\nstream\n" + form + "\nendstream\nendobj"
            }, rootObjectNumber: 1);
        }

        private static string DecodeAllStreams(byte[] pdf) {
            Dictionary<int, PdfIndirectObject> objects = PdfReadDocument.Open(pdf).Objects;
            return string.Concat(objects.Values.Select(static item => item.Value).OfType<PdfStream>()
                .Select(stream => PdfEncoding.Latin1GetString(OfficeIMO.Pdf.Filters.StreamDecoder.DecodeRequired(stream, objects, int.MaxValue))));
        }

        private byte[] BuildLayeredSource(string content) => BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /OFF [6 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /Properties << /Layer 6 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(content)),
            "6 0 obj\n<< /Type /OCG /Name (Hidden layer) >>\nendobj"
        }, rootObjectNumber: 1);

        private byte[] BuildScanWithTextLayer(out string content) {
            // An uncompressed grey image stands in for a scanned page; the OCR layer is invisible text over it.
            const int width = 306, height = 396;
            byte[] pixels = Enumerable.Repeat((byte)0xF0, width * height).ToArray();
            content = "q 612 0 0 792 0 0 cm /Im1 Do Q BT 3 Tr /F1 12 Tf 13.8 TL 72 720 Td (Alpha secret Omega) Tj T* (Gypsy next line) Tj ET";
            return BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Im1 6 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
                BuildStreamObject(5, System.Text.Encoding.ASCII.GetBytes(content)),
                "6 0 obj\n<< /Type /XObject /Subtype /Image /Width " + width + " /Height " + height + " /ColorSpace /DeviceGray /BitsPerComponent 8 /Length " + pixels.Length +
                    " >>\nstream\n" + PdfEncoding.Latin1GetString(pixels) + "\nendstream\nendobj"
            }, rootObjectNumber: 1);
        }

        private static IReadOnlyList<(string Text, double X, double Y)> ReadAllGlyphs(byte[] bytes) =>
            PdfReadDocument.Open(bytes).Pages[0].GetGlyphTextSpans(includeHiddenOptionalContent: true).SelectMany(span => {
                Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
                return glyphs.Select(glyph => (glyph.Text, Math.Round(glyph.X, 3), Math.Round(glyph.Y, 3)));
            }).ToArray();

        // The kept glyphs must be exactly the source glyphs minus the matched characters.
        private static void AssertRemovedExactly(byte[] before, byte[] after, string literal, int expectedOccurrences = 1) {
            List<(string Text, double X, double Y)> expected = ReadAllGlyphs(before).ToList();
            IReadOnlyList<(string Text, double X, double Y)> actual = ReadAllGlyphs(after);
            string lost = string.Concat(expected.Except(actual).Select(static glyph => glyph.Text));
            Assert.Equal(string.Concat(Enumerable.Repeat(literal, expectedOccurrences)).Length, expected.Count - actual.Count);
            Assert.Equal(literal.OrderBy(static c => c).Distinct(), lost.OrderBy(static c => c).Distinct());
            Assert.Empty(actual.Except(expected));
        }
    }
}
