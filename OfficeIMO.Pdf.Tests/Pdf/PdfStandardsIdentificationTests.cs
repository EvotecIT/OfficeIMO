using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStandardsIdentificationTests {
    [Fact]
    public void CombinedArchivalAccessibilityReadbackRejectsMissingExtensionProperty() {
        PdfOptions options = FontOptions().UsePdfA(PdfComplianceProfile.PdfA2B).UsePdfUa();
        byte[] pdf = Document(options).ToBytes();
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA2B, pdf).Requirements
            .Single(r => r.Id == "readback-pdfua-extension-schema").Status);
        // Same byte length preserves xref and stream framing: only the schema declaration is damaged.
        byte[] marker = Encoding.ASCII.GetBytes("<pdfaProperty:name>part</pdfaProperty:name>");
        int at = Find(pdf, marker);
        Assert.True(at >= 0);
        Encoding.ASCII.GetBytes("xxxx").CopyTo(pdf, at + "<pdfaProperty:name>".Length);
        PdfComplianceRequirement requirement = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA2B, pdf)
            .Requirements.Single(r => r.Id == "readback-pdfua-extension-schema");
        Assert.Equal(PdfComplianceRequirementStatus.Missing, requirement.Status);
    }

    [Fact]
    public void UseHelpersEnableStrictGateAndValidateBothClaims() {
        PdfOptions options = FontOptions().UsePdfA(PdfComplianceProfile.PdfA2B).UsePdfUa();
        Assert.Equal(PdfComplianceProfile.PdfA2B, options.ComplianceProfile);
        Assert.Throws<InvalidOperationException>(() => PdfDocument.Create(options).Paragraph(p => p.Text("No UA title")).ToBytes());
        Assert.NotEmpty(Document(options).ToBytes());
        PdfOptions groundwork = FontOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA2B).ConfigurePdfUaGroundwork();
        Assert.Equal(PdfComplianceProfile.None, groundwork.ComplianceProfile);
    }

    [Fact]
    public void PdfA4StoresMetadataOnlyInXmpAndDeclaresRevision() {
        byte[] pdf = Document(FontOptions().UsePdfA(PdfComplianceProfile.PdfA4)).ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);
        Assert.DoesNotContain(" /Info ", raw.Substring(raw.LastIndexOf("trailer", StringComparison.Ordinal)), StringComparison.Ordinal);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        Assert.Equal("Standards sample", info.XmpMetadata!.Title);
        Assert.Contains("<pdfaid:rev>2020</pdfaid:rev>", info.XmpMetadata.RawXml, StringComparison.Ordinal);
    }

    [Fact]
    public void ConflictingPdfVersionsInCombinedClaimsFailClosed() {
        PdfOptions options = FontOptions().UsePdfA(PdfComplianceProfile.PdfA4).UsePdfUa(PdfComplianceProfile.PdfUa1);
        Assert.Throws<InvalidOperationException>(() => Document(options).ToBytes());
    }

    private static PdfOptions FontOptions() => new PdfOptions().UseFontFamily("Carlito", PdfComplianceTestFonts.FindBundledTrueTypeFont()!);
    private static PdfDocument Document(PdfOptions options) => PdfDocument.Create(options)
        .Meta(title: "Standards sample", author: "OfficeIMO")
        .Paragraph(p => p.Text("Accessible archival text: café, Zażółć."));
    private static int Find(byte[] data, byte[] marker) {
        for (int i = 0; i <= data.Length - marker.Length; i++) {
            int j = 0;
            while (j < marker.Length && data[i + j] == marker[j]) j++;
            if (j == marker.Length) return i;
        }
        return -1;
    }
}
