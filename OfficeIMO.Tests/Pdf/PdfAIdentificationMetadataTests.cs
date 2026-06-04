using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAIdentificationMetadataTests {
    [Fact]
    public void PdfAIdentification_CanBeEmittedInXmpWithoutFormalComplianceProfile() {
        var options = new PdfOptions()
            .SetPdfAIdentification(3, "b");

        byte[] bytes = PdfDocument.Create(options)
            .PdfAIdentification(3, "B")
            .Meta(title: "PDF/A identification primitive", author: "OfficeIMO")
            .Paragraph(p => p.Text("PDF/A identification metadata is groundwork, not certification."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfAIdentification cloneIdentification = options.Clone().PdfAIdentification!;

        Assert.True(options.IncludeXmpMetadata);
        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.True(info.HasXmpMetadata);
        Assert.Contains("xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\"", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:part>3</pdfaid:part>", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:conformance>B</pdfaid:conformance>", raw, StringComparison.Ordinal);
        Assert.Equal(3, cloneIdentification.Part);
        Assert.Equal("B", cloneIdentification.Conformance);
    }

    [Fact]
    public void PdfAIdentification_PropertyEmitsXmpEvenWhenGeneralXmpFlagIsFalse() {
        var options = new PdfOptions {
            PdfAIdentification = new PdfAIdentification(2, "U")
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("PDF/A identification property proof."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);

        Assert.False(options.IncludeXmpMetadata);
        Assert.Contains("/Metadata", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:part>2</pdfaid:part>", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:conformance>U</pdfaid:conformance>", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfAIdentification_ValidatesSupportedPartsAndConformanceLevels() {
        Assert.Equal("A", new PdfAIdentification(2, "a").Conformance);
        Assert.Equal("B", new PdfAIdentification(3, "B").Conformance);
        Assert.Equal("U", new PdfAIdentification(3, " u ").Conformance);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfAIdentification(1, "B"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfAIdentification(4, "B"));
        Assert.Throws<ArgumentException>(() => new PdfAIdentification(3, "X"));
        Assert.Throws<ArgumentException>(() => new PdfAIdentification(3, ""));
    }

    [Fact]
    public void PdfAGroundworkHelper_EmitsProfileSpecificArchivalPrerequisitesWithoutFormalProfile() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3A, "pl-PL");

        byte[] bytes = PdfDocument.Create()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3A, "en-GB")
            .Meta(title: "PDF/A groundwork bundle", author: "OfficeIMO")
            .Paragraph(p => p.Text("PDF/A groundwork bundle keeps formal profile generation disabled."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfOptions clone = options.Clone();

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.Equal(PdfFileVersion.Pdf17, clone.FileVersion);
        Assert.True(clone.IncludeXmpMetadata);
        Assert.True(clone.IncludeStandardFontToUnicodeMaps);
        Assert.Equal(3, clone.PdfAIdentification!.Part);
        Assert.Equal("A", clone.PdfAIdentification.Conformance);
        Assert.Equal(PdfOutputIntentPolicy.SrgbIec6196621, clone.OutputIntent!.Policy);
        Assert.Equal(PdfTaggedStructureMode.CatalogMarkers, clone.TaggedStructureMode);
        Assert.Equal("pl-PL", clone.Language);
        Assert.StartsWith("%PDF-1.7", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:part>3</pdfaid:part>", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:conformance>A</pdfaid:conformance>", raw, StringComparison.Ordinal);
        Assert.Contains("/OutputIntents", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Contains("/Lang <656E2D4742>", raw, StringComparison.Ordinal);
        Assert.Contains("/MarkInfo << /Marked true >>", raw, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", raw, StringComparison.Ordinal);
        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasOutputIntents);
        Assert.True(info.HasTaggedContent);
        Assert.Equal("en-GB", info.CatalogLanguage);
        Assert.Throws<ArgumentException>(() => new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfUa1));
    }

    [Fact]
    public void ElectronicInvoiceMetadata_CanBeEmittedInXmpWithoutFormalComplianceProfile() {
        var options = new PdfOptions()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"));

        byte[] bytes = PdfDocument.Create(options)
            .ElectronicInvoiceMetadata("BASIC")
            .Meta(title: "E-invoice XMP primitive", author: "OfficeIMO")
            .Paragraph(p => p.Text("E-invoice metadata is groundwork, not certification."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfElectronicInvoiceMetadata cloneMetadata = options.Clone().ElectronicInvoiceMetadata!;

        Assert.True(options.IncludeXmpMetadata);
        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.True(info.HasXmpMetadata);
        Assert.Contains("xmlns:fx=\"urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#\"", raw, StringComparison.Ordinal);
        Assert.Contains("xmlns:pdfaExtension=\"http://www.aiim.org/pdfa/ns/extension/\"", raw, StringComparison.Ordinal);
        Assert.Contains("<fx:DocumentType>INVOICE</fx:DocumentType>", raw, StringComparison.Ordinal);
        Assert.Contains("<fx:DocumentFileName>factur-x.xml</fx:DocumentFileName>", raw, StringComparison.Ordinal);
        Assert.Contains("<fx:Version>1.0</fx:Version>", raw, StringComparison.Ordinal);
        Assert.Contains("<fx:ConformanceLevel>BASIC</fx:ConformanceLevel>", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaSchema:schema>Factur-X PDF/A Extension Schema</pdfaSchema:schema>", raw, StringComparison.Ordinal);
        Assert.Equal("EN 16931", cloneMetadata.ConformanceLevel);
    }

    [Fact]
    public void FacturXInvoiceXmlHelper_EmitsCanonicalAttachmentAndMatchingXmp() {
        byte[] invoiceXml = CreateCiiXml();
        var options = new PdfOptions()
            .AddFacturXInvoiceXml(invoiceXml, "BASIC", relationship: PdfAssociatedFileRelationship.Alternative);

        byte[] bytes = PdfDocument.Create()
            .AttachFacturXInvoiceXml(invoiceXml, "EN 16931", relationship: PdfAssociatedFileRelationship.Data)
            .Meta(title: "E-invoice attachment primitive", author: "OfficeIMO")
            .Paragraph(p => p.Text("Factur-X/ZUGFeRD attachment metadata is groundwork, not certification."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfExtractedAttachment attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(bytes));
        PdfEmbeddedFile optionAttachment = Assert.Single(options.EmbeddedFiles);
        PdfElectronicInvoiceMetadata cloneMetadata = options.Clone().ElectronicInvoiceMetadata!;

        Assert.True(options.IncludeXmpMetadata);
        Assert.Equal("factur-x.xml", optionAttachment.FileName);
        Assert.Equal("application/xml", optionAttachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Alternative, optionAttachment.Relationship);
        Assert.Equal("BASIC", cloneMetadata.ConformanceLevel);
        Assert.Equal("factur-x.xml", cloneMetadata.DocumentFileName);
        Assert.Contains("<fx:DocumentFileName>factur-x.xml</fx:DocumentFileName>", raw, StringComparison.Ordinal);
        Assert.Contains("<fx:ConformanceLevel>EN 16931</fx:ConformanceLevel>", raw, StringComparison.Ordinal);
        Assert.Contains("/EmbeddedFiles", raw, StringComparison.Ordinal);
        Assert.Contains("/AFRelationship /Data", raw, StringComparison.Ordinal);
        Assert.Equal("factur-x.xml", attachment.Name);
        Assert.Equal("factur-x.xml", attachment.FileName);
        Assert.Equal("factur-x.xml", attachment.UnicodeFileName);
        Assert.Equal("Factur-X/ZUGFeRD invoice XML", attachment.Description);
        Assert.Equal("application/xml", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.Equal(invoiceXml, attachment.Bytes);
    }

    [Fact]
    public void FacturXGroundworkHelper_EmitsPdfA3EinvoicePrerequisitesWithoutFormalProfile() {
        byte[] invoiceXml = CreateCiiXml();
        var options = new PdfOptions()
            .ConfigureFacturXGroundwork(invoiceXml, "EXTENDED");

        byte[] bytes = PdfDocument.Create()
            .ConfigureFacturXGroundwork(invoiceXml, "EN 16931")
            .Meta(title: "E-invoice groundwork bundle", author: "OfficeIMO")
            .Paragraph(p => p.Text("Factur-X/ZUGFeRD groundwork bundle keeps formal profile generation disabled."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfExtractedAttachment attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(bytes));
        PdfOptions clone = options.Clone();

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.Equal(PdfFileVersion.Pdf17, clone.FileVersion);
        Assert.True(clone.IncludeStandardFontToUnicodeMaps);
        Assert.Equal(3, clone.PdfAIdentification!.Part);
        Assert.Equal("B", clone.PdfAIdentification.Conformance);
        Assert.Equal("EXTENDED", clone.ElectronicInvoiceMetadata!.ConformanceLevel);
        Assert.Equal(PdfOutputIntentPolicy.SrgbIec6196621, clone.OutputIntent!.Policy);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, clone.OutputIntent.OutputConditionIdentifier);
        Assert.StartsWith("%PDF-1.7", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:part>3</pdfaid:part>", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:conformance>B</pdfaid:conformance>", raw, StringComparison.Ordinal);
        Assert.Contains("<fx:ConformanceLevel>EN 16931</fx:ConformanceLevel>", raw, StringComparison.Ordinal);
        Assert.Contains("/OutputIntents", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Equal("factur-x.xml", attachment.FileName);
        Assert.Equal("application/xml", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.Equal(invoiceXml, attachment.Bytes);
    }

    [Fact]
    public void PdfUaIdentification_CanBeEmittedInXmpWithoutFormalComplianceProfile() {
        var options = new PdfOptions()
            .SetPdfUaIdentification(PdfUaIdentification.PdfUa1());

        byte[] bytes = PdfDocument.Create(options)
            .PdfUaIdentification()
            .Meta(title: "PDF/UA identification primitive", author: "OfficeIMO")
            .Paragraph(p => p.Text("PDF/UA identification metadata is groundwork, not certification."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfUaIdentification cloneIdentification = options.Clone().PdfUaIdentification!;

        Assert.True(options.IncludeXmpMetadata);
        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.True(info.HasXmpMetadata);
        Assert.Contains("xmlns:pdfuaid=\"http://www.aiim.org/pdfua/ns/id/\"", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfuaid:part>1</pdfuaid:part>", raw, StringComparison.Ordinal);
        Assert.Equal(1, cloneIdentification.Part);
    }

    [Fact]
    public void PdfUaGroundworkHelper_EmitsConfigurableAccessibilityPrerequisitesWithoutFormalProfile() {
        var options = new PdfOptions {
            ViewerPreferences = new PdfViewerPreferencesOptions {
                HideToolbar = true
            }
        }.ConfigurePdfUaGroundwork("pl-PL");

        byte[] bytes = PdfDocument.Create()
            .ConfigurePdfUaGroundwork("en-GB")
            .Meta(title: "PDF/UA groundwork bundle", author: "OfficeIMO")
            .Paragraph(p => p.Text("PDF/UA groundwork bundle keeps formal profile generation disabled."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfOptions clone = options.Clone();

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.Equal(PdfFileVersion.Pdf17, clone.FileVersion);
        Assert.True(clone.IncludeStandardFontToUnicodeMaps);
        Assert.Equal(1, clone.PdfUaIdentification!.Part);
        Assert.Equal(PdfTaggedStructureMode.CatalogMarkers, clone.TaggedStructureMode);
        Assert.Equal("pl-PL", clone.Language);
        Assert.True(clone.ViewerPreferences!.DisplayDocTitle);
        Assert.True(clone.ViewerPreferences.HideToolbar);
        Assert.StartsWith("%PDF-1.7", raw, StringComparison.Ordinal);
        Assert.Contains("<pdfuaid:part>1</pdfuaid:part>", raw, StringComparison.Ordinal);
        Assert.Contains("/Lang <656E2D4742>", raw, StringComparison.Ordinal);
        Assert.Contains("/ViewerPreferences", raw, StringComparison.Ordinal);
        Assert.Contains("/DisplayDocTitle true", raw, StringComparison.Ordinal);
        Assert.Contains("/MarkInfo << /Marked true >>", raw, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasTaggedContent);
        Assert.Equal("en-GB", info.CatalogLanguage);
        Assert.True(info.ViewerPreferences!.GetBoolean("DisplayDocTitle"));
    }

    [Fact]
    public void PdfUaIdentification_ValidatesSupportedPartsAndSnapshotsState() {
        var identification = PdfUaIdentification.PdfUa1();
        var options = new PdfOptions { PdfUaIdentification = identification };

        Assert.Equal(1, options.PdfUaIdentification!.Part);
        Assert.Equal(1, options.Clone().PdfUaIdentification!.Part);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfUaIdentification(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfUaIdentification(2));
    }

    [Fact]
    public void TaggedPdfCatalogMarkers_CanBeEmittedWithoutFormalComplianceProfile() {
        var options = new PdfOptions()
            .EnableTaggedPdfCatalogMarkers();

        byte[] bytes = PdfDocument.Create(options)
            .TaggedStructure(PdfTaggedStructureMode.CatalogMarkers)
            .Meta(title: "Tagged PDF marker primitive", author: "OfficeIMO")
            .Paragraph(p => p.Text("Tagged catalog markers are groundwork, not certification."))
            .ToBytes();

        string raw = Encoding.UTF8.GetString(bytes);
        PdfDocumentProbe probe = PdfInspector.Probe(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfOptions clone = options.Clone();

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.Equal(PdfTaggedStructureMode.CatalogMarkers, clone.TaggedStructureMode);
        Assert.True(probe.HasTaggedContent);
        Assert.True(info.HasTaggedContent);
        Assert.Contains("/MarkInfo << /Marked true >>", raw, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /StructTreeRoot", raw, StringComparison.Ordinal);
        Assert.Contains("/StructParents 0", raw, StringComparison.Ordinal);
        Assert.Contains("/P << /MCID 0 >> BDC", raw, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /P", raw, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedStructureMode_ValidatesSupportedValues() {
        var options = new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        };

        Assert.Equal(PdfTaggedStructureMode.CatalogMarkers, options.TaggedStructureMode);
        Assert.Throws<ArgumentOutOfRangeException>(() => options.TaggedStructureMode = (PdfTaggedStructureMode)99);
    }

    [Fact]
    public void ElectronicInvoiceMetadata_ValidatesAndSnapshotsState() {
        var metadata = PdfElectronicInvoiceMetadata.FacturX("EN 16931");
        var options = new PdfOptions().SetElectronicInvoiceMetadata(metadata);

        metadata.ConformanceLevel = "BASIC";
        PdfElectronicInvoiceMetadata stored = options.ElectronicInvoiceMetadata!;
        stored.ConformanceLevel = "EXTENDED";

        Assert.Equal("EN 16931", options.ElectronicInvoiceMetadata!.ConformanceLevel);
        Assert.Throws<ArgumentException>(() => PdfElectronicInvoiceMetadata.FacturX(""));
        Assert.Throws<ArgumentException>(() => new PdfElectronicInvoiceMetadata("INVOICE", "folder/factur-x.xml", "1.0", "EN 16931"));
        Assert.Throws<ArgumentException>(() => new PdfElectronicInvoiceMetadata("INVOICE", "factur-x.xml", "", "EN 16931"));
        Assert.Throws<ArgumentException>(() => new PdfElectronicInvoiceMetadata("INVOICE", "factur-x.xml", "1.0", " "));
        Assert.Throws<ArgumentException>(() => options.AddFacturXInvoiceXml(CreateCiiXml(), relationship: PdfAssociatedFileRelationship.Supplement));
    }

    [Fact]
    public void FormalComplianceProfilesStillFailClosedWhenPdfAIdentificationIsPresent() {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create(new PdfOptions()
                    .SetPdfAIdentification(3, "B")
                    .RequireCompliance(PdfComplianceProfile.PdfA3B))
                .Paragraph(p => p.Text("Identification alone is not PDF/A-3b support."))
                .ToBytes());

        Assert.Contains("PDF/A-3b", exception.Message, StringComparison.Ordinal);
        Assert.Contains("cannot yet generate certified", exception.Message, StringComparison.Ordinal);
        Assert.Contains("veraPDF validation fixtures in the build lane", exception.Message, StringComparison.Ordinal);
    }

    private static byte[] CreateCiiXml() {
        return Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<rsm:CrossIndustryInvoice xmlns:rsm=\"urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100\" xmlns:ram=\"urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100\" xmlns:udt=\"urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100\">" +
            "<rsm:ExchangedDocumentContext>" +
            "<ram:GuidelineSpecifiedDocumentContextParameter>" +
            "<ram:ID>urn:factur-x.eu:1p0:en16931</ram:ID>" +
            "</ram:GuidelineSpecifiedDocumentContextParameter>" +
            "</rsm:ExchangedDocumentContext>" +
            "<rsm:ExchangedDocument>" +
            "<ram:ID>INV-2026-0001</ram:ID>" +
            "<ram:TypeCode>380</ram:TypeCode>" +
            "<ram:IssueDateTime><udt:DateTimeString format=\"102\">20260603</udt:DateTimeString></ram:IssueDateTime>" +
            "</rsm:ExchangedDocument>" +
            "<rsm:SupplyChainTradeTransaction>" +
            "<ram:IncludedSupplyChainTradeLineItem>" +
            "<ram:AssociatedDocumentLineDocument><ram:LineID>1</ram:LineID></ram:AssociatedDocumentLineDocument>" +
            "<ram:SpecifiedTradeProduct><ram:Name>OfficeIMO PDF compliance work</ram:Name></ram:SpecifiedTradeProduct>" +
            "<ram:SpecifiedLineTradeAgreement>" +
            "<ram:NetPriceProductTradePrice>" +
            "<ram:ChargeAmount currencyID=\"EUR\">100.00</ram:ChargeAmount>" +
            "</ram:NetPriceProductTradePrice>" +
            "</ram:SpecifiedLineTradeAgreement>" +
            "<ram:SpecifiedLineTradeDelivery><ram:BilledQuantity unitCode=\"C62\">1</ram:BilledQuantity></ram:SpecifiedLineTradeDelivery>" +
            "<ram:SpecifiedLineTradeSettlement>" +
            "<ram:ApplicableTradeTax>" +
            "<ram:TypeCode>VAT</ram:TypeCode>" +
            "<ram:CategoryCode>S</ram:CategoryCode>" +
            "<ram:RateApplicablePercent>23</ram:RateApplicablePercent>" +
            "</ram:ApplicableTradeTax>" +
            "<ram:SpecifiedTradeSettlementLineMonetarySummation>" +
            "<ram:LineTotalAmount currencyID=\"EUR\">100.00</ram:LineTotalAmount>" +
            "</ram:SpecifiedTradeSettlementLineMonetarySummation>" +
            "</ram:SpecifiedLineTradeSettlement>" +
            "</ram:IncludedSupplyChainTradeLineItem>" +
            "<ram:ApplicableHeaderTradeAgreement>" +
            "<ram:SellerTradeParty>" +
            "<ram:Name>OfficeIMO Seller</ram:Name>" +
            "<ram:SpecifiedTaxRegistration><ram:ID schemeID=\"VA\">PL1234567890</ram:ID></ram:SpecifiedTaxRegistration>" +
            "<ram:PostalTradeAddress><ram:CountryID>PL</ram:CountryID></ram:PostalTradeAddress>" +
            "</ram:SellerTradeParty>" +
            "<ram:BuyerTradeParty>" +
            "<ram:Name>OfficeIMO Buyer</ram:Name>" +
            "<ram:SpecifiedTaxRegistration><ram:ID schemeID=\"VA\">DE123456789</ram:ID></ram:SpecifiedTaxRegistration>" +
            "<ram:PostalTradeAddress><ram:CountryID>DE</ram:CountryID></ram:PostalTradeAddress>" +
            "</ram:BuyerTradeParty>" +
            "</ram:ApplicableHeaderTradeAgreement>" +
            "<ram:ApplicableHeaderTradeSettlement>" +
            "<ram:InvoiceCurrencyCode>EUR</ram:InvoiceCurrencyCode>" +
            "<ram:ApplicableTradeTax>" +
            "<ram:CalculatedAmount currencyID=\"EUR\">23.45</ram:CalculatedAmount>" +
            "<ram:TypeCode>VAT</ram:TypeCode>" +
            "<ram:BasisAmount currencyID=\"EUR\">100.00</ram:BasisAmount>" +
            "<ram:CategoryCode>S</ram:CategoryCode>" +
            "<ram:RateApplicablePercent>23</ram:RateApplicablePercent>" +
            "</ram:ApplicableTradeTax>" +
            "<ram:SpecifiedTradeSettlementPaymentMeans>" +
            "<ram:TypeCode>58</ram:TypeCode>" +
            "<ram:PayeePartyCreditorFinancialAccount>" +
            "<ram:IBANID>PL61109010140000071219812874</ram:IBANID>" +
            "</ram:PayeePartyCreditorFinancialAccount>" +
            "</ram:SpecifiedTradeSettlementPaymentMeans>" +
            "<ram:SpecifiedTradePaymentTerms>" +
            "<ram:Description>Due within 30 days</ram:Description>" +
            "<ram:DueDateDateTime><udt:DateTimeString format=\"102\">20260703</udt:DateTimeString></ram:DueDateDateTime>" +
            "</ram:SpecifiedTradePaymentTerms>" +
            "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
            "<ram:TaxBasisTotalAmount currencyID=\"EUR\">100.00</ram:TaxBasisTotalAmount>" +
            "<ram:TaxTotalAmount currencyID=\"EUR\">23.45</ram:TaxTotalAmount>" +
            "<ram:GrandTotalAmount currencyID=\"EUR\">123.45</ram:GrandTotalAmount>" +
            "</ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
            "</ram:ApplicableHeaderTradeSettlement>" +
            "</rsm:SupplyChainTradeTransaction>" +
            "</rsm:CrossIndustryInvoice>");
    }
}
