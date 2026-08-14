using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlBogusDoctypeQuotesDoNotHideFollowingElements() {
        AssertHtmlPreflightRejects(
            "<!DOCTYPE html x\"><div></div><div></div>",
            maximumEntries: 1);
    }

    [Fact]
    public void VisioSignatureRelationshipPartHonorsTheXmlNodeBudget() {
        string relationships = string.Concat(Enumerable.Range(0, 32).Select(index =>
            "<Relationship Id=\"rId" + index + "\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature\" Target=\"sig1.xml\"/>"));
        byte[] package = ReplaceWave38Entry(
            CreateSignedVisioProvenancePackage(0),
            "_xmlsignatures/_rels/origin.sigs.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            relationships + "</Relationships>");
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxContainerEntries = 16;

        Assert.Throws<InvalidDataException>(() =>
            VisioDocument.RemoveProvenance(package, options: options));
    }
}
