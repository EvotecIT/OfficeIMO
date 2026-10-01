using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void EpubRootfilePathsAllowPercentEncodedReservedCharactersWithinSegments() {
        string container = Wave63ContainerPrefix +
            "<rootfile full-path='OPS/package%23name.opf' media-type='application/oebps-package+xml'/>" +
            Wave63ContainerSuffix;
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("OPS/package#name.opf", Encoding.UTF8.GetBytes(Wave63Opf)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = EpubDocument.RemoveProvenance(package);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }


}
