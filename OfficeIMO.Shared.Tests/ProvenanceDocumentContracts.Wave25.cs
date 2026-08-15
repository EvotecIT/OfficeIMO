using System.IO.Packaging;
using System.Text;
using OfficeIMO;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightCountsOnlyAsciiTagOpeners() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<é><é><é><é><é><html><head><script type=\"application/c2pa\">" +
            manifest + "</script></head><body></body></html>";
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 8 };

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html, options);

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void HtmlIgnoresImageFallbacksInsideInactiveRules() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html =
            $"<html><head><style>@media print{{.hero{{background:var(--missing,url({dataUri}))}}}}</style></head>" +
            "<body><div class=\"hero\"></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlFileRemovalEncodesDirectlyIntoTheDetectedLegacyCharset() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><meta charset=\"windows-1252\"><script type=\"application/c2pa\">" +
            manifest + "</script></head><body>" + new string('é', 12000) + "</body></html>";
        byte[] input = windows1252.GetBytes(html);
        string inputPath = Path.GetTempFileName();
        string outputPath = Path.GetTempFileName();
        try {
            File.WriteAllBytes(inputPath, input);
            var options = new OfficeProvenanceRemovalOptions();
            options.Limits.MaxAssetBytes = input.Length + 4096;
            options.Limits.MaxManifestBytes = 1024;
            options.Limits.MaxExpandedContainerBytes = input.Length + 4096;

            OfficeProvenanceRemovalResult result = HtmlProvenance.RemoveFile(inputPath, outputPath, options);
            byte[] output = File.ReadAllBytes(outputPath);

            Assert.True(result.WasChanged);
            Assert.True(output.LongLength <= options.Limits.MaxAssetBytes);
            Assert.Contains(new string('é', 12000), windows1252.GetString(output), StringComparison.Ordinal);
        } finally {
            File.Delete(inputPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void VisioSignatureCleanupHonorsTheAppMetadataXmlNodeBudget() {
        byte[] package = CreateSignedVisioProvenancePackageWithAppElements(32);
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxContainerEntries = 16;

        Assert.Throws<InvalidDataException>(() =>
            VisioDocument.RemoveProvenance(package, "drawing.vsdx", options));
    }

    private static byte[] CreateSignedVisioProvenancePackageWithAppElements(int elementCount) {
        byte[] package = CreateSignedVisioProvenancePackage(0);
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        stream.Position = 0;
        using (Package opened = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite)) {
            Uri appUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
            PackagePart app = opened.GetPart(appUri);
            using var writer = new StreamWriter(
                app.GetStream(FileMode.Create, FileAccess.Write),
                new UTF8Encoding(false),
                4096,
                leaveOpen: false);
            writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
            for (int index = 0; index < elementCount; index++) writer.Write($"<Item Index=\"{index}\"/>");
            writer.Write("<DigSig>signature</DigSig></Properties>");
        }
        return stream.ToArray();
    }
}
