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
}
