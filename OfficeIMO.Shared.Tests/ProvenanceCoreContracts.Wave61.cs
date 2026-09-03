using System.Text;
using System.Threading;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData((byte)0x00)]
    [InlineData((byte)0x06)]
    public void WebpExtendedHeaderFlagsMustMatchTheActualFeatureChunks(byte flags) {
        byte[] header = new byte[10];
        header[0] = flags;
        byte[] webp = CreateWebp(
            CreateRiffChunk("VP8X", header),
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateRiffChunk("XMP ", CreateXmpPacket()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ZipSignatureInspectionSharesTheOuterExpandedByteBudget() {
        const string contentTypes =
            "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
            "<Override PartName='/_xmlsignatures/sig1.xml' ContentType='application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml'/>" +
            "</Types>";
        byte[] signature = Encoding.UTF8.GetBytes("<Signature>" + new string('x', 1_200) + "</Signature>");
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("_xmlsignatures/sig1.xml", signature),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxExpandedContainerBytes = 1_490;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(package, "signed.docx", options));

        Assert.Contains("expanded", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RemovalResultHashObservesCancellationBeforeProcessingOwnedBytes() {
        var report = new OfficeProvenanceReport(
            OfficeProvenanceAssetFormat.Png,
            Array.Empty<OfficeProvenanceEvidence>());
        var result = new OfficeProvenanceRemovalResult(
            new byte[1024 * 1024],
            report,
            report,
            Array.Empty<OfficeProvenanceChange>(),
            wasReserialized: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => result.ComputeDataSha256(cancellation.Token));
    }

    [Fact]
    public void OpcSignatureInspectionAcceptsEquivalentDuplicateDefaultContentTypes() {
        const string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Default Extension=\"XML\" ContentType=\"application/xml\"/>" +
            "</Types>";
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("ppt/presentation.xml", Encoding.UTF8.GetBytes("<presentation/>")));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package);

        Assert.False(info.HasSignatures);
    }

    [Fact]
    public void OpcSignatureInspectionRejectsConflictingDuplicateDefaultContentTypes() {
        const string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Default Extension=\"XML\" ContentType=\"text/xml\"/>" +
            "</Types>";
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("ppt/presentation.xml", Encoding.UTF8.GetBytes("<presentation/>")));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficePackageSignatureService.Inspect(package));

        Assert.Contains("conflicting default content types", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
