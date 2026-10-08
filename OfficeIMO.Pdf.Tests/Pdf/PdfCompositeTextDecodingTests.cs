using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCompositeTextDecodingTests {
    [Fact]
    public void IndependentCompositeFontWithoutUnicodeMappingRejectsTextExtraction() {
        byte[] source = LoadIndependentSource();

        NotSupportedException error = Assert.ThrowsAny<NotSupportedException>(
            () => PdfReadDocument.Open(source).ExtractText());

        Assert.Contains("UniJIS-UCS2-H", error.Message);
    }

    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public void IndependentCompositeFontSearchReportsBlockedMappingInsteadOfNoMatches(PdfRedactionTextSelection selection) {
        PdfDocument document = PdfDocument.Load(LoadIndependentSource());

        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions {
            TextSelection = selection
        }.AddLiteral("private account 123"));

        Assert.False(plan.IsReviewable);
        Assert.Empty(plan.Areas);
        PdfDiagnosticFinding finding = Assert.Single(plan.Findings,
            static finding => finding.Code == "RedactionSearchTextMappingUnsupported");
        Assert.Equal(PdfDiagnosticSeverity.Error, finding.Severity);
        Assert.Contains("UniJIS-UCS2-H", finding.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IndependentCompositeFontBlocksManualAndVerificationPlanning(bool verification) {
        byte[] source = LoadIndependentSource();
        PdfRedactionArea[] areas = new[] { new PdfRedactionArea(1, 50, 680, 200, 40) };

        PdfRedactionPlan plan = verification
            ? PdfRedactionPlanner.PlanForVerification(source, areas, options: null)
            : PdfDocument.Load(source).Redactions.Plan(areas);

        Assert.False(plan.IsReviewable);
        Assert.Contains(plan.Findings, static finding => finding.Severity == PdfDiagnosticSeverity.Error &&
            finding.Message.Contains("UniJIS-UCS2-H"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdvertisedCompositeUnicodeMapMustContainMappings(bool emptyMap) {
        ToUnicodeCMap? map = null;
        if (emptyMap) {
            Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes("begincmap endcmap"), out map));
        }
        PdfFontResource font = new PdfFontResource("F1", "HeiseiMin-W3", "UniJIS-UCS2-H", true, map, fontSubtype: "Type0");
        Func<byte[], int, string> decode = ResourceResolver.CreateBudgetedDecoder(font);

        Assert.ThrowsAny<NotSupportedException>(() => decode(new byte[] { 0, 65 }, 10));
        Assert.Equal(string.Empty, decode(Array.Empty<byte>(), 10));
    }

    [Fact]
    public void NamedCompositeFontWithUnicodeMappingDecodesMappedText() {
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes("1 beginbfchar <0041> <0041> endbfchar"), out ToUnicodeCMap? map));
        PdfFontResource font = new PdfFontResource("F1", "HeiseiMin-W3", "UniJIS-UCS2-H", true, map, fontSubtype: "Type0");

        Assert.Equal("A", ResourceResolver.CreateBudgetedDecoder(font)(new byte[] { 0, 65 }, 10));
    }

    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public void PartialCompositeMapCannotProduceAReviewableNoMatchesPlan(PdfRedactionTextSelection selection) {
        byte[] source = BuildNamedCompositeSource("", includePartialMap: true);

        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(
            new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("あい"));

        Assert.False(plan.IsReviewable);
        Assert.Empty(plan.Areas);
    }

    [Fact]
    public void PartialCompositeMapRejectsUnmappedExtractionCodes() {
        byte[] source = BuildNamedCompositeSource("", includePartialMap: true);

        Assert.ThrowsAny<NotSupportedException>(() => PdfReadDocument.Open(source).ExtractText());
    }

    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public void CompleteCompositeMapRemainsReadableAndSearchable(PdfRedactionTextSelection selection) {
        byte[] source = BuildNamedCompositeSource("", includePartialMap: true, mapAllShownCodes: true);

        Assert.Contains("あい", PdfReadDocument.Open(source).ExtractText());
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(
            new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("あい"));

        Assert.True(plan.IsReviewable);
        Assert.Single(plan.Areas);
    }

    [Theory]
    [InlineData("/Artifact BMC", "Known body")]
    [InlineData("/Span << /ActualText (Semantic phrase) >> BDC", "Semantic phrase")]
    public void LogicalCompositeTextRetainsArtifactAndActualTextContracts(string wrapper, string expected) {
        byte[] source = BuildNamedCompositeSource(wrapper, includePartialMap: false);

        string text = PdfReadDocument.Open(source).ExtractText();

        Assert.Contains("Known body", text);
        Assert.Contains(expected, text);
        Assert.DoesNotContain("0B0D", text);
        if (wrapper == "/Artifact BMC") {
            Assert.ThrowsAny<NotSupportedException>(() => PdfReadDocument.Open(source).Pages[0].GetTextSpans(includeArtifactText: true));
        }
    }

    [Fact]
    public void UnsupportedCompositeRunStopsDecodingAfterWholeRunFailure() {
        PdfFontResource font = new PdfFontResource("F2", "HeiseiMin-W3", "UniJIS-UCS2-H", false, null, fontSubtype: "Type0");
        Func<byte[], int, string> decoder = ResourceResolver.CreateBudgetedDecoder(font);
        int decodes = 0;

        Assert.ThrowsAny<NotSupportedException>(() => TextContentParser.Parse(
            "BT /F2 12 Tf <" + string.Concat(Enumerable.Repeat("3042", 2000)) + "> Tj ET",
            (_, bytes) => { decodes++; return decoder(bytes, 10000); },
            (_, bytes) => bytes.Length * 500D));

        Assert.InRange(decodes, 1, 3);
    }

    [Theory]
    [InlineData("/Artifact BMC")]
    [InlineData("/Span << /ActualText (Semantic phrase) >> BDC")]
    public void UnsupportedLogicalGeometryRemainsCancellable(string wrapper) {
        PdfFontResource font = new PdfFontResource("F2", "HeiseiMin-W3", "UniJIS-UCS2-H", false, null, fontSubtype: "Type0");
        Func<byte[], int, string> decoder = ResourceResolver.CreateBudgetedDecoder(font);
        int widths = 0;
        using System.Threading.CancellationTokenSource cancellation = new System.Threading.CancellationTokenSource();

        Assert.ThrowsAny<OperationCanceledException>(() => TextContentParser.Parse(
            wrapper + " BT /F2 12 Tf <" + string.Concat(Enumerable.Repeat("3042", 2000)) + "> Tj ET EMC",
            (_, bytes) => decoder(bytes, 10000),
            (_, bytes) => { if (++widths == 4) cancellation.Cancel(); return bytes.Length * 500D; },
            cancellationCheck: cancellation.Token.ThrowIfCancellationRequested));

        Assert.InRange(widths, 4, 10);
    }

    [Theory]
    [InlineData("/Artifact BMC", false)]
    [InlineData("/Artifact BMC", true)]
    [InlineData("/Span << /ActualText (Semantic phrase) >> BDC", false)]
    [InlineData("/Span << /ActualText (Semantic phrase) >> BDC", true)]
    public void LogicalMappingExemptionsCannotMakeRedactionPlanningReviewable(string wrapper, bool verification) {
        byte[] source = BuildNamedCompositeSource(wrapper, includePartialMap: false);
        PdfRedactionArea[] areas = new[] { new PdfRedactionArea(1, 50, 650, 250, 60) };

        PdfRedactionPlan plan = verification
            ? PdfRedactionPlanner.PlanForVerification(source, areas, options: null)
            : PdfDocument.Load(source).Redactions.Plan(areas);

        Assert.False(plan.IsReviewable);
        Assert.Contains(plan.Findings, static finding => finding.Code == "RedactionSourceTextMappingUnsupported");
    }

    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public void ActualTextSearchCannotProduceAnUnapplicableReviewablePlan(PdfRedactionTextSelection selection) {
        byte[] source = BuildNamedCompositeSource("/Span << /ActualText (Semantic phrase) >> BDC", includePartialMap: false);

        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(
            new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral("Semantic phrase"));

        Assert.False(plan.IsReviewable);
    }

    private static byte[] BuildNamedCompositeSource(string wrapper, bool includePartialMap, bool mapAllShownCodes = false) {
        string content = "BT /F1 12 Tf 72 700 Td (Known body) Tj ET\n" + wrapper +
            "\nBT /F2 12 Tf 72 680 Td <30423044> Tj ET\n" + (wrapper.Length == 0 ? "" : "EMC\n");
        string map = "/CIDInit /ProcSet findresource begin 12 dict begin begincmap " +
            "/CIDSystemInfo << /Registry (Test) /Ordering (Unicode) /Supplement 0 >> def " +
            "/CMapName /PartialTestMap def /CMapType 2 def " +
            "1 begincodespacerange <0000> <FFFF> endcodespacerange " +
            (mapAllShownCodes ? "2 beginbfchar <3042> <3042> <3044> <3044> endbfchar " : "1 beginbfchar <0041> <0041> endbfchar ") +
            "endcmap CMapName currentdict /CMap defineresource pop end end";
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R /F2 6 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "endstream",
            "<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 /Encoding /UniJIS-UCS2-H /DescendantFonts [7 0 R] " +
                (includePartialMap ? "/ToUnicode 8 0 R " : "") + ">>",
            "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 5 >> /DW 500 >>",
            "<< /Length " + map.Length + " >>\nstream\n" + map + "\nendstream"
        };
        StringBuilder pdf = new StringBuilder("%PDF-1.7\n");
        List<int> offsets = new List<int> { 0 };
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 9\n0000000000 65535 f \n");
        foreach (int offset in offsets.Skip(1)) {
            pdf.Append(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        pdf.Append("trailer\n<< /Size 9 /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    [Fact]
    public void UnusedCompositeFontDoesNotBlockKnownText() {
        const string composite = "/Unused << /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 " +
            "/Encoding /UniJIS-UCS2-H /DescendantFonts [<< /Type /Font /Subtype /CIDFontType0 " +
            "/BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 5 >> >>] >>";
        byte[] original = PdfReaderAndFooterRegressionTests.BuildSingleStreamPdf("BT /F1 12 Tf 72 700 Td (Keep me public) Tj ET");
        byte[] source = Encoding.ASCII.GetBytes(Encoding.ASCII.GetString(original)
            .Replace("/Font << /F1 4 0 R >>", "/Font << /F1 4 0 R " + composite + " >>"));

        Assert.Contains("Keep me public", PdfReadDocument.Open(source).ExtractText());
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("Keep"));

        Assert.True(plan.IsReviewable);
        Assert.Single(plan.Areas);
    }

    private static byte[] LoadIndependentSource() => File.ReadAllBytes(Path.Combine(
        AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "CompositeTextDecoding", "unijis-without-to-unicode.pdf"));
}
