using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
#if NET8_0_OR_GREATER
    [Fact]
    public void SignatureAggregateLimitRejectsDeclaredBytesBeforeExpansion() {
        string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "<Override PartName=\"/_xmlsignatures/sig2.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>";
        string validSignature = "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>";
        string oversizedSignature = "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Object>" +
            new string('x', 8 * 1024 * 1024) + "</Object></Signature>";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(archive, "[Content_Types].xml", contentTypes);
                WriteZipEntry(archive, "_xmlsignatures/sig1.xml", oversizedSignature);
                WriteZipEntry(archive, "_xmlsignatures/sig2.xml", validSignature);
            }
            package = output.ToArray();
        }
        int validBytes = Encoding.UTF8.GetByteCount(validSignature);
        var options = new OfficePackageSignatureInspectionOptions {
            VerifyDigests = false,
            MaxSignatureBytes = Encoding.UTF8.GetByteCount(oversizedSignature) + 1L,
            MaxTotalDigestBytes = validBytes + 1L
        };

        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        OfficePackageSignatureInfo inspection = OfficePackageSignatureService.Inspect(package, options);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;

        Assert.True(allocated < 4L * 1024L * 1024L, $"Inspection allocated {allocated:N0} bytes.");
        OfficePackageSignaturePartInfo rejected = Assert.Single(
            inspection.SignatureParts,
            part => part.Uri.EndsWith("sig1.xml", StringComparison.Ordinal));
        OfficePackageSignaturePartInfo accepted = Assert.Single(
            inspection.SignatureParts,
            part => part.Uri.EndsWith("sig2.xml", StringComparison.Ordinal));
        Assert.Contains("aggregate limit", rejected.ParseError ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("aggregate limit", accepted.ParseError ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }
#endif

    [Fact]
    public void ZipRewriteDropsStalePerEntryZip64Extras() {
        byte[] retainedExtra = { 0xFE, 0xCA, 0x01, 0x00, 0x42 };
        byte[] zip64Extra = new byte[20];
        WriteLittleEndian16(zip64Extra, 0, 0x0001);
        WriteLittleEndian16(zip64Extra, 2, 16);
        WriteLittleEndian64(zip64Extra, 4, 4);
        WriteLittleEndian64(zip64Extra, 12, 4);
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream manifest = archive.CreateEntry("META-INF/content_credential.c2pa").Open()) {
                    WriteAll(manifest, CreateManifestStore());
                }
                using (Stream keep = archive.CreateEntry("keep.txt").Open()) {
                    WriteAll(keep, Encoding.UTF8.GetBytes("keep"));
                }
            }
            package = AddEntryExtraFields(
                output.ToArray(),
                "keep.txt",
                Join(retainedExtra, zip64Extra),
                Join(retainedExtra, zip64Extra));
        }

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "keep.txt");

        Assert.Equal(retainedExtra, ReadLocalExtraField(result.ToArray(), centralHeader));
        Assert.Equal(retainedExtra, ReadCentralExtraField(result.ToArray(), centralHeader));
        Assert.Equal("keep", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "keep.txt")));
    }
}
