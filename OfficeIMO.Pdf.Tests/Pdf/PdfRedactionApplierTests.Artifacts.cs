using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    public static IEnumerable<object[]> ArtifactRedactionCases() {
        foreach (PdfRedactionTextSelection selection in new[] { PdfRedactionTextSelection.LogicalBlocks, PdfRedactionTextSelection.MatchedGlyphs })
            foreach (bool hybrid in new[] { false, true })
                foreach (string scope in new[] { "BodyAndArtifacts", "ArtifactsOnly", "PageScoped" })
                    yield return new object[] { selection, hybrid, scope };
    }

    [Theory]
    [MemberData(nameof(ArtifactRedactionCases))]
    public void Search_RemovesPaintedArtifactsRegardlessOfReadingDefault(PdfRedactionTextSelection selection, bool hybrid, string scope) {
        byte[] input = BuildArtifactRedactionFixture(hybrid, scope != "ArtifactsOnly", scope == "ArtifactsOnly");
        using PdfDocument source = PdfDocument.Load(input, new PdfLoadOptions { IncludeArtifactText = false });
        Assert.DoesNotContain("header", source.Reader.Text(), StringComparison.Ordinal);
        var search = new PdfRedactionSearchOptions { TextSelection = selection };
        search.AddLiteral("Word");
        if (scope == "PageScoped") search.PageNumbers.Add(1);
        PdfRedactionPlan plan = source.Redactions.Search(search);
        Assert.True(plan.IsReviewable);
        Assert.NotEmpty(plan.Areas);
        Assert.Contains(plan.Areas, static area => area.PageNumber == 1 && area.Y > 700);
        PdfRedactionApplyResult applied = source.Redactions.ApplyWithEvidence(plan,
            new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.None });
        Assert.True(applied.Evidence.IsVerified);
        Assert.Empty(applied.Evidence.ResidualMatches);
        byte[] output = applied.Pdf;
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.DoesNotContain("Word", independent.GetPage(1).Text, StringComparison.OrdinalIgnoreCase);
        if (selection == PdfRedactionTextSelection.MatchedGlyphs) Assert.Contains("1", independent.GetPage(1).Text, StringComparison.Ordinal);
        if (scope == "PageScoped") Assert.Contains("Word", independent.GetPage(2).Text, StringComparison.Ordinal);
        else {
            Assert.DoesNotContain("Word", independent.GetPage(2).Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Word", WithoutPdfNames(Encoding.GetEncoding(28591).GetString(output)), StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Word", WithoutPdfNames(DecodeSavedResidueBytes(output)), StringComparison.OrdinalIgnoreCase);
        }
        // A structural role remains syntax even when its name equals the search criterion.
        Assert.Contains("/S /Word", DecodeSavedResidueBytes(output), StringComparison.Ordinal);
        foreach (PdfExternalValidator validator in new[] { PdfExternalValidator.QpdfCheck(), PdfExternalValidator.QpdfDecoded(), PdfExternalValidator.PdfText() }) {
            if (!validator.IsAvailable) continue;
            PdfExternalProcessResult external = validator.Run(output, "artifact-redaction.pdf");
            Assert.Equal(0, external.ExitCode);
            if (scope != "PageScoped" && (validator.Name == "PDF text" || external.Output.Contains("%PDF-", StringComparison.Ordinal)))
                Assert.DoesNotContain("Word", WithoutPdfNames(external.Output), StringComparison.OrdinalIgnoreCase);
        }
        if (scope == "ArtifactsOnly") {
            Assert.False(source.Redactions.Verify(new PdfRedactionVerificationOptions {
                CheckRawPdfBytes = false, CheckEncodedPdfStrings = false, CheckDecodedPdfStreams = false
            }.RequireRemovedText("Word")).IsVerified);
            // Manual geometry must also inspect artifacts, and verification must reject an unapplied area.
            PdfRedactionPlan manual = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 735, 220, 35) });
            Assert.Contains(manual.Matches, static match => match.Kind == PdfRedactionMatchKind.TextBlock);
            Assert.False(source.Redactions.VerifyAppliedPlan(manual, new PdfRedactionVerificationOptions()).IsVerified);
            byte[] manualOutput = source.Redactions.Apply(manual.Areas).ToBytes();
            using var manualReader = UglyToad.PdfPig.PdfDocument.Open(manualOutput);
            Assert.DoesNotContain("Word header", manualReader.GetPage(1).Text, StringComparison.Ordinal);
        }
    }

    private byte[] BuildArtifactRedactionFixture(bool hybrid, bool body, bool encodedArtifacts) {
        string resources = "/Resources << /Font << /F1 4 0 R >> /XObject << /Header 8 0 R >> >>";
        var objects = new Dictionary<int, string> {
            [1] = "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 20 0 R >>",
            [2] = "<< /Type /Pages /Count 2 /Kids [3 0 R 6 0 R] /MediaBox [0 0 612 792] >>",
            [3] = "<< /Type /Page /Parent 2 0 R /Contents 5 0 R /StructParents 0 " + resources + " >>",
            [4] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            [5] = Stream(Page(1)),
            [6] = "<< /Type /Page /Parent 2 0 R /Contents 7 0 R /StructParents 1 " + resources + " >>",
            [7] = Stream(Page(2)),
            [8] = Stream("/Artifact BMC BT /F1 12 Tf 72 710 Td <576F7264> Tj ( form header) Tj ET EMC",
                "/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >>"),
            [20] = "<< /Type /StructTreeRoot /K [29 0 R] /ParentTree << /Nums [0 [29 0 R] 1 [29 0 R]] >> >>",
            [29] = "<< /Type /StructElem /S /Word /P 20 0 R /Pg 3 0 R /K 0 >>"
        };
        if (encodedArtifacts) {
            objects[4] = objects[4].Replace("/Encoding /WinAnsiEncoding", "/Encoding /WinAnsiEncoding /ToUnicode 9 0 R");
            // The query exists only through the font mapping, not as a literal stream operand.
            objects[9] = Stream("/CIDInit /ProcSet findresource begin 12 dict begin begincmap " +
                "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def /CMapName /FixtureMap def /CMapType 2 def " +
                "1 begincodespacerange <00> <FF> endcodespacerange 4 beginbfchar " +
                "<41> <0057> <42> <006F> <43> <0072> <44> <0064> endbfchar " +
                "endcmap CMapName currentdict /CMap defineresource pop end end");
            objects[5] = Stream(Page(1).Replace("Word", "ABCD"));
            objects[7] = Stream(Page(2).Replace("Word", "ABCD"));
            objects[8] = objects[8].Replace("576F7264", "41424344");
        }
        return hybrid ? BuildWordHybridRedactionFixture(false, objects) : BuildPdf(objects.Select(item =>
            item.Key.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + item.Value + "\nendobj").ToArray(), 1);

        string Page(int page) => "/P << /MCID 0 >> BDC BT /F1 12 Tf 72 600 Td (" + (body ? "Word body" : "Public body") + ") Tj ET EMC " +
            "/Artifact << /Type /Pagination /Subtype /Header >> BDC BT /F1 12 Tf 72 750 Td (Word header) Tj ET EMC " +
            "/Artifact << /Type /Pagination /Subtype /Footer >> BDC BT /F1 12 Tf 72 30 Td (Word footer ) Tj /Span BMC (" + page + ") Tj EMC ET EMC /Header Do";
        static string Stream(string content, string dictionary = "") => "<< " + dictionary + " /Length " + Encoding.ASCII.GetByteCount(content) + " >>\nstream\n" + content + "\nendstream";
    }

    private static string WithoutPdfNames(string text) => Regex.Replace(text, @"/[^\s()<>\[\]{}/%]+", "");
}
