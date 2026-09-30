using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {


    [Fact]
    public void PowerPointProvenanceAcceptsMacroEnabledAddIns() {
        byte[] package = CreateSavedOpenXmlPackage("pptx");
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        stream.Position = 0;
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
            document.ChangeDocumentType(PresentationDocumentType.AddIn);
        }

        OfficeProvenanceRemovalResult result = PowerPointPresentation.RemoveProvenance(stream.ToArray(), "addin.ppam");

        Assert.Equal(OfficeProvenanceAssetFormat.ZipPackage, result.Before.Format);
    }
}
