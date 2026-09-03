using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ExcelXlsbRejectsMalformedPercentEscapesInWorkbookTargets() {
        byte[] package = RenameWave71Entry(
            CreateWave33XlsbProvenancePackage(
                signed: false,
                officeDocumentTarget: "xl/workbook%ZZ.bin"),
            "xl/workbook.bin",
            "xl/workbook%ZZ.bin");

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }

    [Fact]
    public void ExcelXlsbRequiresDirectRootRelationships() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "_rels/.rels",
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
            "<Extension><Relationship Id='rId1' " +
            "Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' " +
            "Target='xl/workbook.bin'/></Extension></Relationships>");

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }

    [Theory]
    [InlineData("content.xml")]
    [InlineData("CONTENT.XML")]
    public void OdfRejectsDuplicateAndCaseAmbiguousEntries(string duplicateName) {
        byte[] package = CreateZipPackage(
            "odt",
            "META-INF/documentsignatures.xml",
            CreatePngWithManifest(CreateManifestStore()));
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry duplicate = archive.CreateEntry(duplicateName, CompressionLevel.Optimal);
            using Stream target = duplicate.Open();
            byte[] content = Encoding.UTF8.GetBytes("<duplicate/>");
            target.Write(content, 0, content.Length);
        }

        Assert.Throws<InvalidDataException>(() =>
            OdfDocument.RemoveProvenance(stream.ToArray(), "document.odt"));
    }

    [Fact]
    public void SignatureRemovalSharesThePackageRewriteExpansionBudget() {
        byte[] package = CreateZipPackage(
            "odt",
            "META-INF/documentsignatures.xml",
            CreatePngWithManifest(CreateManifestStore()));
        OfficeProvenanceRemovalResult preview = OdfDocument.RemoveProvenance(
            package,
            "document.odt",
            new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.PreserveSignatureMarkup
            });
        long combinedExpandedBytes = GetWave72ExpandedBytes(package) + GetWave72ExpandedBytes(preview.ToArray());
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxExpandedContainerBytes = combinedExpandedBytes - 1;

        Assert.Throws<InvalidDataException>(() =>
            OdfDocument.RemoveProvenance(package, "document.odt", options));
    }

    [Fact]
    public void SignatureRemovalAppliesTheOutputLimitToTheFinalStrippedPackage() {
        byte[] package = CreateZipPackage(
            "odt",
            "META-INF/documentsignatures.xml",
            CreatePngWithManifest(CreateManifestStore()));
        var random = new Random(42);
        var signatureBytes = new byte[16 * 1024];
        random.NextBytes(signatureBytes);
        package = ReplaceWave38Entry(
            package,
            "META-INF/documentsignatures.xml",
            Convert.ToBase64String(signatureBytes));

        OfficeProvenanceRemovalResult preview = OdfDocument.RemoveProvenance(
            package,
            "document.odt",
            new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.PreserveSignatureMarkup
            });
        OfficeProvenanceRemovalResult baseline = OdfDocument.RemoveProvenance(
            package,
            "document.odt",
            new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            });
        long finalSize = baseline.ToArray().LongLength;
        Assert.True(preview.ToArray().LongLength > finalSize);
        var bounded = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures,
            MaxOutputBytes = finalSize
        };
        bounded.Limits.MaxAssetBytes = 128 * 1024;
        bounded.Limits.MaxManifestBytes = bounded.Limits.MaxAssetBytes;

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt", bounded);

        Assert.Equal(finalSize, result.ToArray().LongLength);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    [InlineData("vsdx")]
    public void OpenXmlSignatureRemovalAllowsAnIntermediateAboveTheFinalLimit(string extension) {
        byte[] package = CreateWave73SignedOpenXmlPackage(extension);
        OfficeProvenanceRemovalResult preview = RemoveWave73OpenXml(
            package,
            extension,
            new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.PreserveSignatureMarkup
            });
        OfficeProvenanceRemovalResult baseline = RemoveWave73OpenXml(
            package,
            extension,
            new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            });
        long finalSize = baseline.ToArray().LongLength;
        Assert.True(preview.ToArray().LongLength > finalSize);
        var bounded = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures,
            MaxOutputBytes = finalSize
        };
        bounded.Limits.MaxAssetBytes = 512 * 1024;
        bounded.Limits.MaxManifestBytes = bounded.Limits.MaxAssetBytes;

        OfficeProvenanceRemovalResult result = RemoveWave73OpenXml(package, extension, bounded);

        Assert.Equal(finalSize, result.ToArray().LongLength);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
    }

    private static long GetWave72ExpandedBytes(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.Entries.Sum(entry => entry.Length);
    }

    private static byte[] CreateWave73SignedOpenXmlPackage(string extension) {
        byte[] package;
        if (extension == "vsdx") {
            package = CreateSignedVisioProvenancePackage(0);
        } else {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
            try {
                CreateOpenXmlPackage(path, extension);
                using (Package container = Package.Open(path, FileMode.Open, FileAccess.ReadWrite)) {
                    Uri manifestUri = PackUriHelper.CreatePartUri(new Uri("/META-INF/content_credential.c2pa", UriKind.Relative));
                    using (Stream target = container.CreatePart(manifestUri, "application/c2pa", CompressionOption.Maximum).GetStream()) {
                        byte[] manifest = CreateManifestStore();
                        target.Write(manifest, 0, manifest.Length);
                    }
                    Uri originUri = PackUriHelper.CreatePartUri(new Uri("/_xmlsignatures/origin.sigs", UriKind.Relative));
                    PackagePart origin = container.CreatePart(
                        originUri,
                        "application/vnd.openxmlformats-package.digital-signature-origin",
                        CompressionOption.Maximum);
                    container.CreateRelationship(
                        originUri,
                        TargetMode.Internal,
                        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin");
                    Uri signatureUri = PackUriHelper.CreatePartUri(new Uri("/_xmlsignatures/sig1.xml", UriKind.Relative));
                    _ = container.CreatePart(
                        signatureUri,
                        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                        CompressionOption.Maximum);
                    origin.CreateRelationship(
                        signatureUri,
                        TargetMode.Internal,
                        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature");
                }
                package = File.ReadAllBytes(path);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        var random = new Random(73);
        var payload = new byte[32 * 1024];
        random.NextBytes(payload);
        byte[] signatureXml = Encoding.UTF8.GetBytes(
            "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Object>" +
            Convert.ToBase64String(payload) + "</Object></Signature>");
        using var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        output.Position = 0;
        using (Package container = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
            Uri signatureUri = PackUriHelper.CreatePartUri(new Uri("/_xmlsignatures/sig1.xml", UriKind.Relative));
            using Stream target = container.GetPart(signatureUri).GetStream(FileMode.Create, FileAccess.Write);
            target.Write(signatureXml, 0, signatureXml.Length);
        }
        return output.ToArray();
    }

    private static OfficeProvenanceRemovalResult RemoveWave73OpenXml(
        byte[] package,
        string extension,
        OfficeProvenanceRemovalOptions options) => extension switch {
            "docx" => WordDocument.RemoveProvenance(package, "document.docx", options),
            "xlsx" => ExcelDocument.RemoveProvenance(package, "workbook.xlsx", options),
            "pptx" => PowerPointPresentation.RemoveProvenance(package, "presentation.pptx", options),
            "vsdx" => VisioDocument.RemoveProvenance(package, "drawing.vsdx", options),
            _ => throw new ArgumentOutOfRangeException(nameof(extension))
        };
}
