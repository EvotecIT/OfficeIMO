using System.IO.Compression;
using System.Text;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void RejectedSignaturePartDoesNotConsumeTheAggregateReadBudget() {
        string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "<Override PartName=\"/_xmlsignatures/sig2.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>";
        string validSignature = "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>";
        string oversizedSignature = "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Object>" +
            new string('x', 512) + "</Object></Signature>";
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
            MaxSignatureBytes = validBytes + 1L,
            MaxTotalDigestBytes = validBytes + 1L
        };

        OfficePackageSignatureInfo inspection = OfficePackageSignatureService.Inspect(package, options);

        OfficePackageSignaturePartInfo rejected = Assert.Single(inspection.SignatureParts, part => part.Uri.EndsWith("sig1.xml", StringComparison.Ordinal));
        OfficePackageSignaturePartInfo accepted = Assert.Single(inspection.SignatureParts, part => part.Uri.EndsWith("sig2.xml", StringComparison.Ordinal));
        Assert.True(rejected.HasParseError);
        Assert.DoesNotContain("aggregate limit", accepted.ParseError ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }
}
