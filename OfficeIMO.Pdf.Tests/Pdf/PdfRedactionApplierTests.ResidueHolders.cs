using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    private const string ResidueHolderQuery = "TOPSECRET";

    [Fact]
    public void CompleteXrefProofAcceptsTablesLongerThanTheRecoveryCap() {
        const int objectCount = 110_000;
        var objects = new List<string>(objectCount) {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj"
        };
        for (int number = 4; number <= objectCount; number++) objects.Add(number + " 0 obj\nnull\nendobj");
        byte[] pdf = BuildPdf(objects, rootObjectNumber: 1);

        Assert.True(PdfSyntax.IsClassicRedactionOutput(pdf, PdfReadDocument.Open(pdf), default));
    }

    [Fact]
    [Trait("Category", "Performance")]
    public void ResidueSweepStaysLinearInStreamCount() {
        const int streamCount = 20_000;
        var objects = new List<string>(streamCount + 5) {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj"
        };
        byte[] data = Encoding.ASCII.GetBytes("plain note");
        for (int number = 4; number < 4 + streamCount; number++) objects.Add(BuildStreamObject(number, data));
        PdfDocument source = PdfDocument.Load(BuildPdf(objects, rootObjectNumber: 1));

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral(ResidueHolderQuery));
        stopwatch.Stop();

        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10), "Residue sweep took " + stopwatch.Elapsed + ".");
    }

    [Theory]
    [InlineData("begincmap note TOPSECRET here", false)]
    [InlineData("begincmap 1 beginbfchar <41> <0041> endbfchar TOPSECRET endcmap", true)]
    public void WordBegincmapDoesNotHideStreamFromResidueSweep(string stream, bool referencedAsToUnicode) {
        string font = referencedAsToUnicode ? "/ToUnicode 6 0 R" : string.Empty;
        PdfDocument source = PdfDocument.Load(BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /PrivateData 6 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica " + font + " >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Public page) Tj ET")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(stream))
        }, rootObjectNumber: 1));

        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral(ResidueHolderQuery));

        Assert.Contains(plan.Findings, static finding => finding.Code == "redaction-search-unplanned-residue");
    }
}
