using System.Xml;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceProfileTests {
    [Theory]
    [InlineData(InvoiceProfile.Minimum, "MINIMUM")]
    [InlineData(InvoiceProfile.BasicWithoutLines, "BASIC WL")]
    [InlineData(InvoiceProfile.Basic, "BASIC")]
    [InlineData(InvoiceProfile.En16931, "EN 16931")]
    [InlineData(InvoiceProfile.Extended, "EXTENDED")]
    [InlineData(InvoiceProfile.XRechnung, "XRECHNUNG")]
    public void CanonicalCiiDeclarationsResolveToMatchingXmp(InvoiceProfile profile, string xmp) {
        var declaration = InvoiceProfileDeclaration.Read(Cii(InvoiceProfiles.GetGuidelineId(profile)));
        Assert.Equal(InvoiceSyntax.Cii, declaration.Syntax);
        Assert.Equal(profile, declaration.Profile);
        Assert.Equal(xmp, InvoiceProfiles.GetXmpConformanceLevel(declaration.Profile!.Value));
    }

    [Theory]
    [InlineData("urn:factur-x.eu:1p0:en16931")]
    [InlineData("urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0-unverified")]
    [InlineData("urn:ferd:CrossIndustryDocument:invoice:1p0:comfort")]
    public void UnsupportedGuidelinesAreInspectableButNotRecognized(string guideline) {
        var declaration = InvoiceProfileDeclaration.Read(Cii(guideline));
        Assert.Equal(guideline, declaration.GuidelineId);
        Assert.Null(declaration.Profile);
    }

    [Fact]
    public void UblPeppolDeclarationCannotBeAdvertisedAsFacturX() {
        string xml = "<Invoice xmlns='urn:oasis:names:specification:ubl:schema:xsd:Invoice-2' xmlns:c='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'><c:CustomizationID>" +
            InvoiceProfiles.PeppolBis30Guideline + "</c:CustomizationID></Invoice>";
        var declaration = InvoiceProfileDeclaration.Read(Encoding.UTF8.GetBytes(xml));
        Assert.Equal(InvoiceSyntax.Ubl, declaration.Syntax);
        Assert.Equal(InvoiceProfile.PeppolBis, declaration.Profile);
        Assert.Throws<ArgumentException>(() => InvoiceProfiles.GetXmpConformanceLevel(declaration.Profile!.Value));
    }

    [Fact]
    public void ProfileInspectionRejectsDuplicateAndNamespaceSpoofedFields() {
        string xml = Encoding.UTF8.GetString(Cii(InvoiceProfiles.En16931Guideline));
        string duplicate = xml.Replace("</a:ID>", "</a:ID><a:ID>urn:factur-x.eu:1p0:basic</a:ID>");
        Assert.Throws<InvalidDataException>(() => InvoiceProfileDeclaration.Read(Encoding.UTF8.GetBytes(duplicate)));
        string spoof = xml.Replace("urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100", "urn:spoof");
        Assert.Throws<InvalidDataException>(() => InvoiceProfileDeclaration.Read(Encoding.UTF8.GetBytes(spoof)));
        Assert.Throws<XmlException>(() => InvoiceProfileDeclaration.Read(Encoding.UTF8.GetBytes(xml + "<trailing/>")));
    }

    [Fact]
    public void ProfileInspectionDoesNotExpandDtdOrAcceptExcessiveDepth() {
        Assert.Throws<XmlException>(() => InvoiceProfileDeclaration.Read(Encoding.UTF8.GetBytes("<!DOCTYPE invoice [<!ENTITY a 'x'>]><invoice>&a;</invoice>")));
        Assert.Throws<InvalidDataException>(() => InvoiceProfileDeclaration.Read(new byte[InvoiceProfileDeclaration.MaximumXmlBytes + 1]));
        string deep = string.Concat(Enumerable.Repeat("<a>", 130)) + string.Concat(Enumerable.Repeat("</a>", 130));
        Assert.Throws<InvalidDataException>(() => InvoiceProfileDeclaration.Read(Encoding.UTF8.GetBytes(deep)));
    }

    internal static byte[] Cii(string guideline) => Encoding.UTF8.GetBytes(
        "<r:CrossIndustryInvoice xmlns:r='urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100' xmlns:a='urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100'>" +
        "<r:ExchangedDocumentContext><a:GuidelineSpecifiedDocumentContextParameter><a:ID>" + guideline +
        "</a:ID></a:GuidelineSpecifiedDocumentContextParameter></r:ExchangedDocumentContext></r:CrossIndustryInvoice>");
}
