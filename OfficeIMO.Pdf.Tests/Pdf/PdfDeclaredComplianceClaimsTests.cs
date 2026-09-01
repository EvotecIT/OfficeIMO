using OfficeIMO.Pdf;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDeclaredComplianceClaimsTests {
    [Fact]
    public void DeclaredClaims_DoNotTreatMetadataAsConformanceProof() {
        PdfComplianceArtifact artifact = PdfDocument.Create(
                new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B))
            .Meta(title: "Declared compliance gate")
            .CreateComplianceArtifact(PdfComplianceProfile.PdfA3B);
        byte[] bytes = artifact.ToBytes();

        PdfDeclaredComplianceClaimsReport report = PdfDocument.Load(bytes).AssessDeclaredComplianceClaims();
        PdfDeclaredComplianceClaim claim = Assert.Single(report.Claims);

        Assert.Equal(PdfArtifactFingerprint.ComputeSha256(bytes), report.ArtifactSha256);
        Assert.Equal(bytes.LongLength, report.ArtifactSizeBytes);
        Assert.Equal(PdfDeclaredComplianceStandard.PdfA, claim.Standard);
        Assert.Equal("PDF/A-3b", claim.Declaration);
        Assert.Equal(PdfComplianceProfile.PdfA3B, claim.Profile);
        Assert.Equal(PdfDeclaredComplianceClaimStatus.MissingExternalValidation, claim.Status);
        Assert.False(claim.CanClaimConformance);
        Assert.False(report.CanClaimAllDeclaredConformance);
    }

    [Fact]
    public void DeclaredClaims_AcceptArtifactBoundExternalValidationOnlyAfterInternalReadiness() {
        PdfComplianceArtifact artifact = PdfDocument.Create(
                new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B))
            .Meta(title: "Exact declared compliance proof")
            .CreateComplianceArtifact(PdfComplianceProfile.PdfA3B);
        byte[] bytes = artifact.ToBytes();
        PdfExternalValidationResult validation = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "PDF/A-3b profile accepted.",
            bytes,
            "PDF/A-3b");

        PdfDeclaredComplianceClaimsReport report = PdfDocument.Load(bytes).AssessDeclaredComplianceClaims(new[] { validation });
        PdfDeclaredComplianceClaim claim = Assert.Single(report.Claims);

        Assert.Equal(PdfDeclaredComplianceClaimStatus.Claimable, claim.Status);
        Assert.True(claim.CanClaimConformance);
        Assert.True(report.CanClaimAllDeclaredConformance);
        Assert.True(claim.Proof!.HasRequiredExternalValidation);
    }

    [Fact]
    public void DeclaredClaims_ReportUnsupportedPdfAOneWithoutInventingAProfile() {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "openpreserve-pdfa1b-text.pdf");

        PdfDeclaredComplianceClaim claim = Assert.Single(PdfDocument.Load(path).AssessDeclaredComplianceClaims().Claims);

        Assert.Equal("PDF/A-1b", claim.Declaration);
        Assert.Null(claim.Profile);
        Assert.False(claim.IsRecognized);
        Assert.Equal(PdfDeclaredComplianceClaimStatus.UnsupportedProfile, claim.Status);
        Assert.Contains("not mapped", claim.Diagnostic, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeclaredClaims_ReturnEmptyReportForOrdinaryPdf() {
        byte[] bytes = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Ordinary PDF.")).ToBytes();

        PdfDeclaredComplianceClaimsReport report = PdfDocument.Load(bytes).AssessDeclaredComplianceClaims();

        Assert.False(report.HasClaims);
        Assert.Empty(report.Claims);
        Assert.False(report.CanClaimAllDeclaredConformance);
    }

    [Fact]
    public void DeclaredClaims_ReadNamespacedRdfAttributeProperties() {
        const string xmp = "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"><rdf:Description rdf:about=\"\" xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\" xmlns:pdfuaid=\"http://www.aiim.org/pdfua/ns/id/\" pdfaid:part=\"1\" pdfaid:conformance=\"B\" pdfuaid:part=\"1\"/></rdf:RDF></x:xmpmeta>";
        byte[] pdf = BuildPdfWithXmp(xmp);

        PdfXmpMetadataInfo metadata = Assert.IsType<PdfXmpMetadataInfo>(PdfReadDocument.Open(pdf).XmpMetadata);
        PdfDeclaredComplianceClaimsReport report = PdfDocument.Load(pdf).AssessDeclaredComplianceClaims();

        Assert.Equal(1, metadata.PdfAPart);
        Assert.Equal("B", metadata.PdfAConformance);
        Assert.Equal(1, metadata.PdfUaPart);
        Assert.Collection(
            report.Claims,
            claim => {
                Assert.Equal(PdfDeclaredComplianceStandard.PdfA, claim.Standard);
                Assert.Equal("PDF/A-1b", claim.Declaration);
                Assert.Equal(PdfDeclaredComplianceClaimStatus.UnsupportedProfile, claim.Status);
            },
            claim => {
                Assert.Equal(PdfDeclaredComplianceStandard.PdfUa, claim.Standard);
                Assert.Equal("PDF/UA-1", claim.Declaration);
                Assert.Equal(PdfComplianceProfile.PdfUa1, claim.Profile);
            });
    }

    private static byte[] BuildPdfWithXmp(string xmp) {
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            StreamObject(string.Empty),
            StreamObject(xmp, "/Type /Metadata /Subtype /XML")
        };
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>(objects.Length);
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xrefOffset = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        for (int index = 0; index < offsets.Count; index++) {
            builder.Append(offsets[index].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Length + 1).Append(" >>\nstartxref\n")
            .Append(xrefOffset.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static string StreamObject(string content, string additionalDictionary = "") {
        int length = Encoding.ASCII.GetByteCount(content);
        string suffix = string.IsNullOrWhiteSpace(additionalDictionary) ? string.Empty : " " + additionalDictionary;
        return "<< /Length " + length.ToString(CultureInfo.InvariantCulture) + suffix + " >>\nstream\n" + content + "\nendstream";
    }
}
