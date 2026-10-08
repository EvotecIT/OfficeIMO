using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDocumentInputAdmissionTests {
    [Theory]
    [InlineData("")]
    [InlineData("This is not a PDF document.")]
    [InlineData("%PDF-1.7\nbroken")]
    public void LenientReadRejectsInputWithNoRecoverableObjects(string input) {
        byte[] bytes = Encoding.ASCII.GetBytes(input);

        AssertNoObjects(() => PdfReadDocument.Open(bytes));
        PdfDocument document = PdfDocument.Load(bytes);
        AssertNoObjects(() => document.InspectForViewing());
        AssertNoObjects(() => document.InspectGeometryForViewing());
        AssertNoObjects(() => document.Read());

        PdfDocumentPreflight preflight = document.Preflight();
        Assert.False(preflight.CanRead);
        Assert.True(preflight.HasReadBlocker(input.StartsWith("%PDF-", StringComparison.Ordinal)
            ? PdfReadBlockerKind.ParserUnsupported : PdfReadBlockerKind.MissingHeader));
    }

    [Fact]
    public async Task FileAndStreamSnapshotsUseTheSameReadAdmission() {
        byte[] bytes = Encoding.ASCII.GetBytes("This is not a PDF document.");
        string path = Path.Combine(Path.GetTempPath(), "officeimo-invalid-pdf-" + Guid.NewGuid().ToString("N") + ".pdf");
        File.WriteAllBytes(path, bytes);
        try {
            AssertNoObjects(() => PdfReadDocument.Open(path));
            AssertNoObjects(() => PdfDocument.Load(path).InspectForViewing());
            PdfDocument asyncFile = await PdfDocument.LoadAsync(path);
            AssertNoObjects(() => asyncFile.InspectForViewing());

            using var stream = new MemoryStream(bytes);
            stream.Position = 5;
            AssertNoObjects(() => PdfReadDocument.Open(stream));
            Assert.Equal(5, stream.Position);
            AssertNoObjects(() => PdfDocument.Load(stream).InspectForViewing());
            PdfDocument asyncStream = await PdfDocument.LoadAsync(stream);
            AssertNoObjects(() => asyncStream.InspectGeometryForViewing());
            Assert.True(stream.CanRead);
            Assert.Equal(5, stream.Position);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void ZeroPageCatalogAndForensicObjectsRemainReadable() {
        byte[] zeroPages = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 3 >>\n%%EOF\n");
        byte[] fragment = Encoding.ASCII.GetBytes(
            "1 0 obj\n<< /Producer (Forensic fragment) >>\nendobj\n");

        Assert.Empty(PdfReadDocument.Open(zeroPages).Pages);
        PdfReadDocument forensic = PdfReadDocument.Open(fragment);
        Assert.Single(forensic.Objects);
        Assert.Empty(forensic.Pages);
        Assert.Contains(forensic.RepairReport.Diagnostics, diagnostic => diagnostic.Code == "MissingStartXref");
    }

    [Fact]
    public void GeneratedPreflightUsesOneSerializedSnapshot() {
        int compositionCalls = 0;
        PdfDocument CreateDocument() => PdfDocument.Create(new PdfOptions {
            TextLineBreakCallback = text => {
                compositionCalls++;
                return new[] { text.Length };
            }
        }).Paragraph(paragraph => paragraph.Text(new string('W', 600)));

        CreateDocument().ToBytes();
        int callsForOneSerialization = compositionCalls;
        Assert.True(callsForOneSerialization > 0);
        compositionCalls = 0;

        PdfDocumentPreflight preflight = CreateDocument().Preflight();

        Assert.True(preflight.CanRead);
        Assert.Single(Assert.IsType<PdfDocumentInfo>(preflight.DocumentInfo).Pages);
        Assert.Equal(callsForOneSerialization, compositionCalls);
    }

    [Fact]
    public void RawProbeRetainsEvidenceWhenThereAreNoReadableObjects() {
        byte[] bytes = Encoding.ASCII.GetBytes("%PDF-1.7\nbroken");

        var (objects, _) = PdfSyntax.ParseObjects(bytes);
        Assert.Empty(objects);
        Assert.Equal("1.7", PdfInspector.Probe(bytes).HeaderVersion);
        PdfArtifactSnapshot artifact = PdfArtifactSnapshot.Capture(bytes);
        Assert.Null(artifact.PageCount);
        Assert.Equal(bytes.LongLength, artifact.ByteCount);
        Assert.False(string.IsNullOrEmpty(artifact.Sha256));
    }

    private static void AssertNoObjects(Action read) {
        PdfParseException error = Assert.Throws<PdfParseException>(read);
        Assert.Equal("NoIndirectObjects", error.Code);
        Assert.Null(error.ObjectNumber);
    }
}
