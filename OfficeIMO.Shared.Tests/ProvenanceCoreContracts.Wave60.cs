using System.Text;
using OfficeIMO;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void TrailingJpegContinuationPreservesTheCompleteCarrierUnderStrictRemoval() {
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 7, sequence: 1),
            CreateJpegApp11(manifest, 0, 1, instance: 7, sequence: 2),
            CreateMinimalJpegFrame(),
            CreateMinimalJpegScan(),
            new byte[] { 0, 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void EscapedOpcApplicationMetadataTargetsRemainDiscoverable() {
        const string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/docProps/app%20custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
            "</Types>";
        const string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app%20custom.xml\"/>" +
            "</Relationships>";
        const string properties =
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
            "<DigSig>signature</DigSig></Properties>";
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("_rels/.rels", Encoding.UTF8.GetBytes(relationships)),
            ("docProps/app%20custom.xml", Encoding.UTF8.GetBytes(properties)));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package);

        Assert.True(info.HasApplicationSignatureMetadata);
        Assert.True(info.HasSignatures);
    }

    [Fact]
    public void UndefinedSignatureMutationPoliciesAreRejectedBeforeRemoval() {
        byte[] package = CreateZip(("META-INF/content_credential.c2pa", CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = (OfficeSignatureMutationPolicy)99
        };

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            OfficeProvenanceRemover.Remove(package, "fixture.zip", options));
    }
}
