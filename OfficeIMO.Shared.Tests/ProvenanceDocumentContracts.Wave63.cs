using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    private const string Wave63ContainerPrefix =
        "<container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles>";
    private const string Wave63ContainerSuffix = "</rootfiles></container>";
    private const string Wave63Opf =
        "<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\">" +
        "<metadata><identifier xmlns=\"http://purl.org/dc/elements/1.1/\" id=\"id\">fixture</identifier></metadata>" +
        "<manifest/><spine/></package>";

    [Fact]
    public void EpubRootfileMediaTypeIsAsciiCaseInsensitive() {
        string container = Wave63ContainerPrefix +
            "<rootfile full-path=\"OPS/package.opf\" media-type=\"Application/OEBPS-Package+XML\"/>" +
            Wave63ContainerSuffix;
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("OPS/package.opf", Encoding.UTF8.GetBytes(Wave63Opf)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = EpubDocument.RemoveProvenance(package);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void EpubRejectsBackslashEntryNames() {
        string container = Wave63ContainerPrefix +
            "<rootfile full-path=\"OPS/package.opf\" media-type=\"application/oebps-package+xml\"/>" +
            Wave63ContainerSuffix;
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("OPS/package.opf", Encoding.UTF8.GetBytes(Wave63Opf)),
            ("OPS\\foreign.xml", Encoding.UTF8.GetBytes("<foreign/>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package));
    }

    [Fact]
    public void IgnoredSelectEndTagsDoNotInflateHtmlPreflightCounts() {
        const string html = "<html><body><div><select></div><img><img></select>";
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 5 };

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html, options);

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void PseudoElementCustomPropertiesResolveOnTheSamePseudoElement() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box::before{--img:url('" + dataUri +
            "');background-image:var(--img)}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TranscodedMarkdownChangesDoNotReportUtf8PhysicalByteCounts(bool bigEndian) {
        Encoding encoding = bigEndian ? Encoding.BigEndianUnicode : Encoding.Unicode;
        string markdown = "before\n-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\nafter\n";
        byte[] input = encoding.GetPreamble().Concat(encoding.GetBytes(markdown)).ToArray();
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        try {
            File.WriteAllBytes(inputPath, input);

            OfficeProvenanceRemovalResult result = MarkdownProvenance.RemoveFile(inputPath, outputPath);

            Assert.True(result.WasChanged);
            Assert.NotEmpty(result.Changes);
            Assert.All(result.Changes, change => Assert.Equal(0, change.RemovedBytes));
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }
}
