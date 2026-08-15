using System.IO.Packaging;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ExcelXlsbSignatureCleanupRewritesRelationshipResolvedApplicationMetadata() {
        byte[] package = CreateXlsbWithNoncanonicalApplicationSignatureMetadata();
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options);

        Assert.True(result.WasChanged);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.DoesNotContain(
            "DigSig",
            Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "metadata/application.xml")),
            StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlDataUrisRejectNonAsciiWhitespaceInsideBase64Payloads() {
        string base64 = Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string inertDataUri = "data:image/png;base64," + base64.Insert(base64.Length / 2, "\u00A0");
        string html = $"<html><body><img src=\"{inertDataUri}\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void SrcsetKeepsCommaContainingUrlTokensIntact() {
        string base64 = Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string srcset = "fallback.png,data:image/png;base64," + base64;
        string html = $"<html><body><img srcset=\"{srcset}\"></body></html>";

        HtmlSrcSetCandidate candidate = Assert.Single(HtmlSrcSetParser.Parse(srcset));
        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);

        Assert.Equal(srcset, candidate.Url);
        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void MarkdownFileLimitsApplyToPhysicalEncodingRatherThanInternalUtf8() {
        byte[] input = Encoding.Unicode.GetPreamble()
            .Concat(Encoding.Unicode.GetBytes(new string('\u20AC', 10)))
            .ToArray();
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        try {
            File.WriteAllBytes(inputPath, input);
            var inspectionOptions = new OfficeProvenanceOptions {
                MaxAssetBytes = 24,
                MaxManifestBytes = 16
            };
            var removalOptions = new OfficeProvenanceRemovalOptions();
            removalOptions.Limits.MaxAssetBytes = 24;
            removalOptions.Limits.MaxManifestBytes = 16;

            Assert.Empty(MarkdownProvenance.InspectFile(inputPath, inspectionOptions).Evidence);
            OfficeProvenanceRemovalResult result = MarkdownProvenance.RemoveFile(inputPath, outputPath, removalOptions);
            Assert.False(result.WasChanged);
            Assert.Equal(input, File.ReadAllBytes(outputPath));
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    private static byte[] CreateXlsbWithNoncanonicalApplicationSignatureMetadata() {
        byte[] original = CreateWave33XlsbProvenancePackage(signed: true);
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);
        output.Position = 0;
        using (Package package = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
            const string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
            foreach (PackageRelationship relationship in package.GetRelationshipsByType(relationshipType).ToArray()) {
                package.DeleteRelationship(relationship.Id);
            }
            Uri canonicalUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
            if (package.PartExists(canonicalUri)) package.DeletePart(canonicalUri);
            Uri customUri = PackUriHelper.CreatePartUri(new Uri("/metadata/application.xml", UriKind.Relative));
            PackagePart application = package.CreatePart(
                customUri,
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                CompressionOption.Maximum);
            using (var writer = new StreamWriter(application.GetStream(), new UTF8Encoding(false), 4096, leaveOpen: false)) {
                writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>signature</DigSig></Properties>");
            }
            package.CreateRelationship(customUri, TargetMode.Internal, relationshipType);
        }
        return output.ToArray();
    }
}
