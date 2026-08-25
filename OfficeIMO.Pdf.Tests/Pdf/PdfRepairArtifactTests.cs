using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRepairArtifactTests {
    [Fact]
    public void Create_PersistsRecoveredStructureAndProvesStrictReadback() {
        byte[] source = BuildStreamPdf("/Length 999");

        PdfRepairArtifactResult result = PdfRepairArtifact.Create(source);

        Assert.True(result.IsVerified);
        Assert.Contains(result.SourceRepairReport.Diagnostics, static diagnostic => diagnostic.Code == "IncorrectStreamLength" && diagnostic.WasRecovered);
        Assert.Empty(result.StrictOutputRepairReport.Diagnostics);
        Assert.Contains("Recovered stream text", PdfReadDocument.Open(result.ToBytes(), new PdfReadOptions { ParsingMode = PdfParsingMode.Strict }).ExtractText(), StringComparison.Ordinal);
        Assert.True(result.Preservation.IsPreserved);
    }

    [Fact]
    public void Create_RejectsCleanAndDetectedOnlySourcesByDefault() {
        byte[] clean = PdfProductionWorkflowTestSupport.CreatePdf("Clean artifact");
        byte[] ambiguous = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Names << /Dests 7 0 R >> >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "7 0 obj\n<< /Names [(dangling)] >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 8 >>\nstartxref\n0\n%%EOF\n");

        Assert.Throws<InvalidOperationException>(() => PdfRepairArtifact.Create(clean));
        Assert.Throws<InvalidOperationException>(() => PdfRepairArtifact.Create(ambiguous, new PdfRepairArtifactOptions { RequireRecoveredDefects = false }));
    }

    [Fact]
    public void Create_PreservesTheCallerSourceInputByteLimit() {
        byte[] source = BuildStreamPdf("/Length 999");
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = source.LongLength - 1L }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfRepairArtifact.Create(source, readOptions: readOptions));

        Assert.Equal(PdfReadLimitKind.InputBytes, exception.Kind);
    }

    [Fact]
    public void Create_ReservesCanonicalRewriteObjectGrowth() {
        const string catalogBody = "<</Type/Catalog/Pages 2 0 R/A[]/B[]/C[]/D[]/E[]/F[]/G[]/H[]/I[]/J[]>>";
        byte[] source = BuildStreamPdf("/Length 999", catalogBody);
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxObjectCharacters = catalogBody.Length,
                MaxTokensPerObject = catalogBody.Length
            }
        };

        PdfRepairArtifactResult result = PdfRepairArtifact.Create(source, readOptions: readOptions);

        Assert.True(result.IsVerified);
        Assert.Single(result.ToDocument().Read.Pages());
    }

    private static byte[] BuildStreamPdf(
        string lengthEntry,
        string catalogBody = "<< /Type /Catalog /Pages 2 0 R >>") {
        const string streamData = "BT (Recovered stream text) Tj ET";
        return Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n" + catalogBody + "\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n<< " + lengthEntry + " >>\nstream\n" + streamData + "\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 5 >>\nstartxref\n0\n%%EOF\n");
    }
}
