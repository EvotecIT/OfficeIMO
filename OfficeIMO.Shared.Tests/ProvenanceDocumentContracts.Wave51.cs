using System.IO.Packaging;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlForeignFontBreakoutUsesAttributeNamesNotQuotedValues() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><svg><font title=\" color=x\"><![CDATA[" +
            string.Concat(Enumerable.Repeat("<div></div>", 64)) +
            "]]></font></svg></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 10 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void PictureSourceSrcIsNotAResponsiveImageCandidate() {
        string sourceDataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string fallbackDataUri = "data:image/png;name=fallback;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><body><picture><source src=\"" + sourceDataUri +
            "\"><img src=\"" + fallbackDataUri + "\"></picture></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string rewritten = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("src=\"" + sourceDataUri + "\"", rewritten, StringComparison.Ordinal);
        Assert.DoesNotContain("src=\"" + fallbackDataUri + "\"", rewritten, StringComparison.Ordinal);
    }

    [Fact]
    public void VisioSignatureCleanupIgnoresUnrelatedCanonicalApplicationPart() {
        byte[] package = CreateVisioPackageWithNoncanonicalApplicationSignatureMetadata();
        using (var output = new MemoryStream()) {
            output.Write(package, 0, package.Length);
            output.Position = 0;
            using (Package container = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
                Uri canonicalUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
                PackagePart unrelated = container.CreatePart(canonicalUri, "application/xml", CompressionOption.Maximum);
                using var writer = new StreamWriter(unrelated.GetStream(), new UTF8Encoding(false), 4096, leaveOpen: false);
                writer.Write("<unrelated>keep</unrelated>");
            }
            package = output.ToArray();
        }
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = VisioDocument.RemoveProvenance(package, "drawing.vsdx", options);

        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.Equal("<unrelated>keep</unrelated>", Encoding.UTF8.GetString(
            ReadZipEntry(result.ToArray(), "docProps/app.xml")));
        Assert.DoesNotContain("DigSig", Encoding.UTF8.GetString(
            ReadZipEntry(result.ToArray(), "metadata/application.xml")), StringComparison.Ordinal);
    }
}
