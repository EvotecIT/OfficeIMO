using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
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

    private static long GetWave72ExpandedBytes(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.Entries.Sum(entry => entry.Length);
    }
}
