using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {

    public sealed class PdfPredefinedCjkTests {
        [Fact]
        public void IndependentMismatchedCollectionRecoversReadableTextOnlyForOptedInInspection() {
            byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures",
                "Interoperability", "PredefinedCjk", "mismatched-collection.pdf"));
            Assert.ThrowsAny<NotSupportedException>(() => PdfReadDocument.Open(source).ExtractText());
            PdfDocument pdf = PdfDocument.Load(source);
            PdfInspectionReadResult ordinary = pdf.ReadForInspection();
            Assert.Null(Assert.Single(ordinary.PageOutcomes).Page);
            Assert.Empty(ordinary.Document.Pages);

            var options = new PdfLoadOptions { RecoverDecodableTextForInspection = true };
            PdfInspectionReadResult recovered = pdf.ReadForInspection(options);
            PdfPageReadOutcome outcome = Assert.Single(recovered.PageOutcomes);
            Assert.Equal(PdfPageReadStatus.UnmappableText, outcome.Status);
            Assert.False(outcome.Succeeded);
            Assert.False(recovered.CompletedAllSelectedPages);
            Assert.Same(Assert.Single(recovered.Document.Pages), outcome.Page);
            Assert.Equal("UniGB-UCS2-H", outcome.MappingFailure!.EncodingName);
            Assert.Equal("Public summary remains readable.", recovered.Document.Text.Trim());
            Assert.Contains("Public summary remains readable.", recovered.Document.ExtractMarkdown());
            PdfFontInspectionDiagnostic warning = Assert.Single(recovered.Document.TextDecodingDiagnostics);
            Assert.Equal(PdfFontInspectionDiagnosticCode.UnmappableText, warning.Code);
            Assert.Equal(1, warning.PageNumber);
            Assert.EndsWith("/Font/F2", warning.ResourcePath);
            Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts,
                Assert.Throws<PdfReadLimitException>(() => pdf.ReadForInspection(new PdfLoadOptions {
                    RecoverDecodableTextForInspection = true,
                    Pipeline = new PdfUnderstandingPipelineOptions { MaxRunsPerPage = 1 }
                })).Kind);

            PdfDocument withRecoveryOption = PdfDocument.Load(source, options);
            Assert.Throws<PdfUnsupportedTextMappingException>(() => withRecoveryOption.Read());
            PdfRedactionPlan plan = withRecoveryOption.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("秘密"));
            Assert.False(plan.IsReviewable);
            Assert.Contains(plan.Findings, static finding => finding.Code == "redaction-search-text-mapping-unsupported");
        }

        [Theory]
        [InlineData("japanese", "機密")]
        [InlineData("simplified-chinese", "秘密")]
        [InlineData("traditional-chinese", "秘密")]
        [InlineData("korean", "비밀")]
        public void IndependentPredefinedFontsDecodeMeasureAndRedactWithoutToUnicode(string name, string term) {
            string fixtures = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "PredefinedCjk");
            byte[] original = File.ReadAllBytes(Path.Combine(fixtures, name + ".pdf"));
            using JsonDocument metadata = JsonDocument.Parse(File.ReadAllText(Path.Combine(fixtures, "source.json")));
            JsonElement expected = metadata.RootElement.GetProperty("cases").EnumerateArray().Single(
                item => item.GetProperty("file").GetString() == name + ".pdf");
            PdfReadDocument read = PdfReadDocument.Open(original);
            PdfTextSpan line = Assert.Single(read.Pages[0].GetTextSpans(), span => span.Text.Contains(term));
            Assert.Equal(expected.GetProperty("advancePoints").GetDouble(), line.Advance, 5);
            Assert.Contains(term, read.ExtractText());
            PdfDocument source = PdfDocument.Load(original);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions {
                TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
                ContentScope = PdfRedactionContentScope.TextOnly
            }.AddLiteral(term));
            Assert.True(plan.IsReviewable, string.Join("; ", plan.Findings.Select(static finding => finding.Message)));
            Assert.Single(plan.Areas);
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan).ThrowIfUnverified();
            string text = result.ToDocument().Read().Text;
            Assert.DoesNotContain(term, text);
            Assert.Contains("Before", text);
            Assert.Contains("after page", text);
            Assert.Contains("Public summary remains readable.", text);
            Assert.Equal(original, source.ToBytes());
        }

        [Theory]
        [InlineData("Other", "Japan1")]
        [InlineData("Adobe", "GB1")]
        public void IncompatibleCharacterCollectionsCannotSelectPredefinedMappings(string registry, string ordering) {
            Assert.Null(PdfPredefinedCMap.Find("UniJIS-UCS2-H", registry, ordering));
        }

        [Fact]
        public void PredefinedCodesEnforceDecodedOutputBudgetAndRejectSurrogates() {
            PdfPredefinedCMap mapping = PdfPredefinedCMap.Find("UniJIS-UCS2-H", "Adobe", "Japan1")!.Value;
            Assert.Throws<PdfReadLimitException>(() => mapping.TryDecode(Encoding.BigEndianUnicode.GetBytes("AB"), 1, out _));
            Assert.False(mapping.TryDecode(new byte[] { 0xD8, 0x00 }, 10, out _));
            Assert.False(mapping.TryDecode(new byte[] { 0x30 }, 10, out _));
        }

        [Fact]
        public void AdvertisedUnicodeOverridesPredefinedMapWithoutFallingBackForMissingCodes() {
            Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes("1 beginbfchar <0041> <0058> endbfchar"), out ToUnicodeCMap? unicode));
            PdfFontResource font = new PdfFontResource("F1", "HeiseiMin-W3", "UniJIS-UCS2-H", true, unicode,
                fontSubtype: "Type0", predefinedCMap: PdfPredefinedCMap.Find("UniJIS-UCS2-H", "Adobe", "Japan1"));
            Func<byte[], int, string> decoder = ResourceResolver.CreateBudgetedDecoder(font);

            Assert.Equal("X", decoder(Encoding.BigEndianUnicode.GetBytes("A"), 10));
            Assert.ThrowsAny<NotSupportedException>(() => decoder(Encoding.BigEndianUnicode.GetBytes("AB"), 10));
        }

        [Theory]
        [InlineData("<D800> Tj", true, true)]
        [InlineData("<30> Tj", true, false)]
        [InlineData("<004130> Tj", true, false)]
        [InlineData("[<00> 10 <41>] TJ", false, false)]
        public void InspectionRecoversAnchoredTextWhenAnInvalidCjkCodeHasNoPaintedWidth(
            string invalidShow, bool strictMappingFails, bool visualCountsUnavailable) {
            byte[] source = BuildInspectionRecoverySource(
                "BT /F1 12 Tf 1 0 0 1 20 170 Tm (Readable before) Tj /F2 12 Tf " + invalidShow + " " +
                "/F1 12 Tf (Position depends on malformed width) Tj 1 0 0 1 20 140 Tm (Readable after anchor) Tj ET " +
                "q 1 0 0 1 20 20 cm /Im0 Do Q", includeImage: true);
            PdfDocument pdf = PdfDocument.Load(source);
            if (strictMappingFails) {
                Assert.Throws<PdfUnsupportedTextMappingException>(() => pdf.Read());
                Assert.Null(Assert.Single(pdf.ReadForInspection().PageOutcomes).Page);
            }

            PdfInspectionReadResult inspection = pdf.ReadForInspection(
                new PdfLoadOptions { RecoverDecodableTextForInspection = true });

            PdfPageReadOutcome outcome = Assert.Single(inspection.PageOutcomes);
            PdfLogicalPage page = Assert.Single(inspection.Document.Pages);
            Assert.Same(page, outcome.Page);
            Assert.Equal(PdfPageReadStatus.UnmappableText, outcome.Status);
            Assert.False(inspection.CompletedAllSelectedPages);
            Assert.Contains("Readable before", inspection.Document.Text);
            Assert.DoesNotContain("Position depends on malformed width", inspection.Document.Text);
            Assert.DoesNotContain(page.TextBlocks.SelectMany(block => block.Spans), span => span.FontResource == "F2");
            PdfLogicalTextBlock reanchored = Assert.Single(page.TextBlocks, block => block.Text == "Readable after anchor");
            Assert.Equal(20D, reanchored.XStart, 5);
            Assert.EndsWith("/Font/F2", Assert.Single(inspection.Document.TextDecodingDiagnostics).ResourcePath);
            Assert.Single(page.Images);
            if (visualCountsUnavailable) Assert.False(page.HasVisualPrimitiveCounts);
        }

        [Fact]
        public void InspectionRetriesOnlyTheUnmappableKernedArrayAsIndividualOperands() {
            byte[] source = BuildInspectionRecoverySource(
                "BT /F2 12 Tf 1 0 0 1 20 170 Tm [15 <0041> 10 <0043> 10 <0042> 15] TJ " +
                "1 0 0 1 20 140 Tm [<0041> 10 <0042>] TJ ET",
                unicode: "2 beginbfchar <0041> <0058> <0042> <0059> endbfchar");
            PdfInspectionReadResult inspection = PdfDocument.Load(source).ReadForInspection(
                new PdfLoadOptions { RecoverDecodableTextForInspection = true });

            PdfPageReadOutcome outcome = Assert.Single(inspection.PageOutcomes);
            Assert.Equal(PdfPageReadStatus.UnmappableText, outcome.Status);
            Assert.False(inspection.CompletedAllSelectedPages);
            PdfTextSpan[] spans = Assert.Single(inspection.Document.Pages).TextBlocks.SelectMany(block => block.Spans).ToArray();
            PdfTextSpan first = Assert.Single(spans, span => span.Text == "X");
            PdfTextSpan second = Assert.Single(spans, span => span.Text == "Y");
            Assert.Equal(19.82D, first.X, 5);
            Assert.True(second.X > first.X);
            Assert.Equal(20D, Assert.Single(spans, span => span.Text == "XY").X, 5);
            Assert.DoesNotContain("C", inspection.Document.Text);
            Assert.Single(inspection.Document.TextDecodingDiagnostics);
        }

        [Theory]
        [InlineData("0009", "1 1 250", 4.5)]
        [InlineData("E000", "0 0 100", 1.8)]
        public void ExplicitUnicodeUsesPaintedCidWidthsIncludingNotdef(string code, string widths, double expectedAdvance) {
            byte[] source = BuildMappedTextSource("UniJIS-UCS2-H", "Japan1", code, widths, code + "> <0041");

            PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

            Assert.Equal("A", span.Text);
            Assert.Equal(expectedAdvance, span.Advance, 5);
        }

        [Theory]
        [InlineData("0 65535 700", 12.6)]
        [InlineData("0 65535 700 4826 [900] 4825 4827 800 4826 [600]", 10.8)]
        [InlineData("4826 [900] 0 65535 700", 12.6)]
        [InlineData("0 65535 700 4825 4827 800", 14.4)]
        [InlineData("4825 4827 800 0 65535 700", 12.6)]
        [InlineData("4825 4827 800 4826 [600]", 10.8)]
        [InlineData("4826 [600] 4825 4827 800", 14.4)]
        public void CidWidthRangesRetainCompleteCoverageAndDeclarationPrecedence(string widths, double expectedAdvance) {
            byte[] source = BuildMappedTextSource("UniCNS-UCS2-H", "CNS1", "6A5F", widths);

            PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

            Assert.Equal("機", span.Text);
            Assert.Equal(expectedAdvance, span.Advance, 5);
        }

        [Fact]
        public void IdentityFontsAlsoUseCompleteCidWidthRanges() {
            byte[] source = BuildMappedTextSource("Identity-H", "Japan1", "12DA", "0 65535 700", "12DA> <0041");

            PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

            Assert.Equal("A", span.Text);
            Assert.Equal(12.6D, span.Advance, 5);
        }

        private static byte[] BuildMappedTextSource(string encoding, string ordering, string code, string widths, string? unicodePair = null) {
            string content = "BT /F1 18 Tf 72 700 Td <" + code + "> Tj ET";
            string unicode = "1 beginbfchar <" + unicodePair + "> endbfchar";
            string[] objects = {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>",
                "<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 /Encoding /" + encoding + " /DescendantFonts [5 0 R] " + (unicodePair == null ? "" : "/ToUnicode 7 0 R ") + ">>",
                "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (" + ordering + ") /Supplement 0 >> /DW 1000 /W [" + widths + "] >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
                "<< /Length " + unicode.Length + " >>\nstream\n" + unicode + "\nendstream"
            };
            return BuildPdf(objects);
        }

        private static byte[] BuildInspectionRecoverySource(string content, string? unicode = null, bool includeImage = false) {
            int imageObject = unicode is null ? 8 : 9;
            string imageResources = includeImage ? " /XObject << /Im0 " + imageObject + " 0 R >>" : string.Empty;
            var objects = new List<string> {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R /F2 5 0 R >>" +
                    imageResources + " >> /Contents 7 0 R >>",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
                "<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 /Encoding /UniJIS-UCS2-H /DescendantFonts [6 0 R]" +
                    (unicode is null ? "" : " /ToUnicode 8 0 R") + " >>",
                "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 0 >> /DW 1000 >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream"
            };
            if (unicode is not null) objects.Add("<< /Length " + unicode.Length + " >>\nstream\n" + unicode + "\nendstream");
            if (includeImage) objects.Add("<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream");
            return BuildPdf(objects);
        }

        private static byte[] BuildPdf(IReadOnlyList<string> objects) {
            StringBuilder pdf = new StringBuilder("%PDF-1.7\n");
            List<int> offsets = new List<int> { 0 };
            for (int index = 0; index < objects.Count; index++) {
                offsets.Add(pdf.Length);
                pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
            }
            int xref = pdf.Length;
            pdf.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
            foreach (int offset in offsets.Skip(1)) pdf.Append(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
            pdf.Append("trailer\n<< /Size ").Append(objects.Count + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
            return Encoding.ASCII.GetBytes(pdf.ToString());
        }
    }
}
