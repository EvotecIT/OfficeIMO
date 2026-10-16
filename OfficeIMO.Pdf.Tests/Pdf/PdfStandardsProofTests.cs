using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

[CollectionDefinition("PDF standards proof output", DisableParallelization = true)]
public sealed class PdfStandardsProofCollection { }

/// <summary>Produces strict-profile artifacts for independent validation in the compliance lane.</summary>
[Collection("PDF standards proof output")]
public sealed class PdfStandardsProofTests {
    [Theory]
    [InlineData("pdfa2a", PdfComplianceProfile.PdfA2A, PdfComplianceProfile.None, 2, "A", null)]
    [InlineData("pdfa2u", PdfComplianceProfile.PdfA2U, PdfComplianceProfile.None, 2, "U", null)]
    [InlineData("pdfa3a", PdfComplianceProfile.PdfA3A, PdfComplianceProfile.None, 3, "A", null)]
    [InlineData("pdfa3u", PdfComplianceProfile.PdfA3U, PdfComplianceProfile.None, 3, "U", null)]
    [InlineData("pdfa4", PdfComplianceProfile.PdfA4, PdfComplianceProfile.None, 4, null, null)]
    [InlineData("pdfa4e", PdfComplianceProfile.PdfA4E, PdfComplianceProfile.None, 4, "E", null)]
    [InlineData("pdfa4f", PdfComplianceProfile.PdfA4F, PdfComplianceProfile.None, 4, "F", null)]
    [InlineData("pdfua2", PdfComplianceProfile.None, PdfComplianceProfile.PdfUa2, null, null, 2)]
    [InlineData("pdfa2b-ua1", PdfComplianceProfile.PdfA2B, PdfComplianceProfile.PdfUa1, 2, "B", 1)]
    [InlineData("pdfa3b-ua1", PdfComplianceProfile.PdfA3B, PdfComplianceProfile.PdfUa1, 3, "B", 1)]
    public void StrictProfileArtifactRetainsEveryDeclaredIdentity(string name, PdfComplianceProfile archival,
        PdfComplianceProfile accessible, int? part, string? conformance, int? uaPart) {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] font = File.ReadAllBytes(fontPath!);
        var options = new PdfOptions { IncludeStandardFontToUnicodeMaps = true };
        if (archival != PdfComplianceProfile.None) options.UsePdfA(archival);
        if (accessible != PdfComplianceProfile.None) options.UsePdfUa(accessible);
        options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
        if (archival == PdfComplianceProfile.PdfA4F) {
            options.AddEmbeddedFile("evidence.txt", Encoding.UTF8.GetBytes("Archival attachment proof"),
                "text/plain", PdfAssociatedFileRelationship.Data);
            options.SetEmbeddedFileModificationDate("evidence.txt", new DateTimeOffset(2026, 10, 16, 0, 0, 0, TimeSpan.Zero));
        }

        byte[] bytes = PdfDocument.Create(options)
            .Meta(title: "OfficeIMO standards proof", author: "OfficeIMO")
            .H1("Standards proof")
            .Paragraph(p => p.Text("Embedded Unicode text and declared profile evidence."))
            .Table(new[] { new[] { "Item", "Amount" }, new[] { "Service", "119.00 EUR" } })
            .ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.Equal(part, xmp.PdfAPart);
        Assert.Equal(conformance, xmp.PdfAConformance);
        Assert.Equal(uaPart, xmp.PdfUaPart);
        Assert.Equal(part == 4 || uaPart == 2 ? "2.0" : "1.7", info.HeaderVersion);
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Embedded Unicode text", text);
        Assert.Contains("Service", text);
        Assert.Contains("119.00 EUR", text);
        Assert.Equal(archival == PdfComplianceProfile.PdfA4F, info.HasEmbeddedFiles);

        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        output = Path.Combine(output, "standards");
        Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, name + ".pdf"), bytes);
    }
}
