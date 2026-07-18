using System;
using System.IO;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComplianceGateTests {
    private const string ProofOutputEnv = "OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT";

    [Fact]
    public void PdfA2BFixture_ContainsArchivalPrimitivesAndEmbeddedFonts() {
        byte[] bytes = CreatePdfA2BFixture();
        WriteProofPdf("officeimo-pdfa2b.pdf", bytes);
        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal("1.7", info.HeaderVersion);
        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasOutputIntents);
        Assert.False(info.HasEmbeddedFiles);
        Assert.Equal("en-US", info.CatalogLanguage);
        Assert.Contains("pdfaid:part>2<", raw, StringComparison.Ordinal);
        Assert.Contains("pdfaid:conformance>B<", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile3 ", raw, StringComparison.Ordinal);
        Assert.Matches(@"/ID \[<[0-9A-F]{32}> <[0-9A-F]{32}>\]", raw);
    }

    [Fact]
    public void VeraPdfGate_AcceptsFormalPdfA2BFixture() {
        byte[] fixture = CreatePdfA2BFixture();
        WriteProofPdf("officeimo-pdfa2b.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf("2b");
        if (!validator.IsAvailable) {
            WriteProofText("verapdf-pdfa2b.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-pdfa2b.pdf");
        WriteProofText("verapdf-pdfa2b.txt", result.GetDiagnosticText());

        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void PdfA3BFixture_ContainsFormalArchivalPrimitivesAndEmbeddedFonts() {
        byte[] bytes = CreatePdfA3BFixture();
        WriteProofPdf("officeimo-pdfa3b.pdf", bytes);
        WriteProfileProofContract();
        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal("1.7", info.HeaderVersion);
        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasOutputIntents);
        Assert.True(info.HasEmbeddedFiles);
        Assert.Equal("en-US", info.CatalogLanguage);
        Assert.Contains("/Metadata ", raw, StringComparison.Ordinal);
        Assert.Contains("/OutputIntents [", raw, StringComparison.Ordinal);
        Assert.Contains("/Length 3052", raw, StringComparison.Ordinal);
        Assert.Contains("pdfaid:part", raw, StringComparison.Ordinal);
        Assert.Contains("pdfaid:conformance", raw, StringComparison.Ordinal);
        Assert.Contains("/EmbeddedFiles", raw, StringComparison.Ordinal);
        Assert.Contains("/AF [", raw, StringComparison.Ordinal);
        Assert.Contains("/Params << /Size 29 /CheckSum <AEEE18719BF2A42A30C88BB9B14D60FE> >>", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile3 ", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void VeraPdfGate_AcceptsFormalPdfA3BFixture() {
        byte[] fixture = CreatePdfA3BFixture();
        WriteProofPdf("officeimo-pdfa3b.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf();
        if (!validator.IsAvailable) {
            WriteProofText("verapdf-pdfa3b.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-pdfa3b.pdf");
        WriteProofText("verapdf-pdfa3b.txt", result.GetDiagnosticText());

        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void MustangGate_AcceptsFormalFacturXFixture() {
        byte[] fixture = CreateElectronicInvoiceFixture(PdfComplianceProfile.FacturX);
        WriteProofPdf("officeimo-facturx.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.Mustang();
        if (!validator.IsAvailable) {
            WriteProofText("mustang-facturx.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-facturx.pdf");
        WriteProofText("mustang-facturx.txt", result.GetDiagnosticText());

        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void VeraPdfGate_AcceptsFormalFacturXFixture() {
        byte[] fixture = CreateElectronicInvoiceFixture(PdfComplianceProfile.FacturX);
        WriteProofPdf("officeimo-facturx.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf();
        if (!validator.IsAvailable) {
            WriteProofText("verapdf-facturx.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-facturx.pdf");
        WriteProofText("verapdf-facturx.txt", result.GetDiagnosticText());

        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void MustangGate_AcceptsFormalZugferdFixture() {
        byte[] fixture = CreateElectronicInvoiceFixture(PdfComplianceProfile.Zugferd);
        WriteProofPdf("officeimo-zugferd.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.Mustang();
        if (!validator.IsAvailable) {
            WriteProofText("mustang-zugferd.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-zugferd.pdf");
        WriteProofText("mustang-zugferd.txt", result.GetDiagnosticText());
        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void VeraPdfGate_AcceptsFormalZugferdFixture() {
        byte[] fixture = CreateElectronicInvoiceFixture(PdfComplianceProfile.Zugferd);
        WriteProofPdf("officeimo-zugferd.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf();
        if (!validator.IsAvailable) {
            WriteProofText("verapdf-zugferd.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-zugferd.pdf");
        WriteProofText("verapdf-zugferd.txt", result.GetDiagnosticText());
        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void PdfUa1Fixture_ContainsFormalAccessibilityPrimitivesAndEmbeddedFonts() {
        byte[] bytes = CreatePdfUa1Fixture();
        WriteProofPdf("officeimo-pdfua1.pdf", bytes);
        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal("1.7", info.HeaderVersion);
        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasTaggedContent);
        Assert.Equal("en-US", info.CatalogLanguage);
        Assert.Contains("pdfuaid:part", raw, StringComparison.Ordinal);
        Assert.Contains("/ViewerPreferences", raw, StringComparison.Ordinal);
        Assert.Contains("/DisplayDocTitle true", raw, StringComparison.Ordinal);
        Assert.Contains("/MarkInfo << /Marked true >>", raw, StringComparison.Ordinal);
        Assert.Contains("/StructTreeRoot", raw, StringComparison.Ordinal);
        Assert.Contains("/S /Annot", raw, StringComparison.Ordinal);
        Assert.Contains("/S /Form", raw, StringComparison.Ordinal);
        Assert.Contains("/FontFile3 ", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfUaValidatorGate_AcceptsFormalPdfUa1Fixture() {
        byte[] fixture = CreatePdfUa1Fixture();
        WriteProofPdf("officeimo-pdfua1.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.PdfUa();
        if (!validator.IsAvailable) {
            WriteProofText("pdfua-pdfua1.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-pdfua1.pdf");
        WriteProofText("pdfua-pdfua1.txt", result.GetDiagnosticText());

        Assert.Equal(0, result.ExitCode);
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfUa2, "PDF/UA identification XMP")]
    public void FormalProfiles_StillFailClosedUntilValidatorBackedGenerationExists(PdfComplianceProfile profile, string expectedRequirement) {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create()
                .Compliance(profile)
                .Meta(title: "Compliance profile guard")
                .Paragraph(p => p.Text("Profile output must fail closed until validator evidence exists."))
                .ToBytes());

        Assert.Contains("cannot yet generate certified", exception.Message, StringComparison.Ordinal);
        Assert.Contains(expectedRequirement, exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ProofOutputHook_WritesComplianceFixturesWhenConfigured() {
        string previousOutput = Environment.GetEnvironmentVariable(ProofOutputEnv) ?? string.Empty;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfComplianceProof", Guid.NewGuid().ToString("N"));
        try {
            Environment.SetEnvironmentVariable(ProofOutputEnv, directory);

            WriteProofPdf("officeimo-pdfa3b.pdf", CreatePdfA3BFixture());
            WriteProofPdf("officeimo-pdfa2b.pdf", CreatePdfA2BFixture());
            WriteProofPdf("officeimo-facturx.pdf", CreateElectronicInvoiceFixture(PdfComplianceProfile.FacturX));
            WriteProofPdf("officeimo-zugferd.pdf", CreateElectronicInvoiceFixture(PdfComplianceProfile.Zugferd));
            WriteProofPdf("officeimo-pdfua1.pdf", CreatePdfUa1Fixture());
            WriteProfileProofContract();
            WriteProofText("verapdf-pdfa3b.txt", "validator diagnostic");
            WriteProofText("verapdf-facturx.txt", "validator diagnostic");
            WriteProofText("verapdf-zugferd.txt", "validator diagnostic");
            WriteProofText("mustang-facturx.txt", "validator diagnostic");
            WriteProofText("mustang-zugferd.txt", "validator diagnostic");
            WriteProofText("pdfua-pdfua1.txt", "validator diagnostic");

            Assert.True(File.Exists(Path.Combine(directory, "officeimo-pdfa3b.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-pdfa2b.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-facturx.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-zugferd.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-pdfua1.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-profile-proof-contract.json")));
            Assert.True(File.Exists(Path.Combine(directory, "verapdf-pdfa3b.txt")));
            Assert.True(File.Exists(Path.Combine(directory, "verapdf-facturx.txt")));
            Assert.True(File.Exists(Path.Combine(directory, "verapdf-zugferd.txt")));
            Assert.True(File.Exists(Path.Combine(directory, "pdfua-pdfua1.txt")));
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-pdfa3b.pdf")).Length > 0);
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-pdfa2b.pdf")).Length > 0);
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-facturx.pdf")).Length > 0);
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-zugferd.pdf")).Length > 0);
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-pdfua1.pdf")).Length > 0);
        } finally {
            Environment.SetEnvironmentVariable(ProofOutputEnv, string.IsNullOrEmpty(previousOutput) ? null : previousOutput);
            TryDeleteDirectory(directory);
        }
    }

    [Fact]
    public void ProofOutputHook_WritesProductProofContractWhenConfigured() {
        string previousOutput = Environment.GetEnvironmentVariable(ProofOutputEnv) ?? string.Empty;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfComplianceProof", Guid.NewGuid().ToString("N"));
        try {
            Environment.SetEnvironmentVariable(ProofOutputEnv, directory);

            WriteProfileProofContract();

            string path = Path.Combine(directory, "officeimo-profile-proof-contract.json");
            Assert.True(File.Exists(path));

            using JsonDocument document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            JsonElement root = document.RootElement;
            Assert.Equal(4, root.GetProperty("schemaVersion").GetInt32());
            Assert.Equal("ExactArtifactValidationInjectedByProofExporter", root.GetProperty("externalEvidenceMode").GetString());
            JsonElement profiles = root.GetProperty("profiles");
            Assert.Equal(5, profiles.GetArrayLength());
            Assert.Contains(profiles.EnumerateArray(), profile =>
                string.Equals(profile.GetProperty("profile").GetString(), nameof(PdfComplianceProfile.PdfA2B), StringComparison.Ordinal) &&
                profile.GetProperty("isInternallyReady").GetBoolean() &&
                profile.GetProperty("hasArtifactEvidence").GetBoolean() &&
                profile.GetProperty("artifactSha256").GetString()!.Length == 64 &&
                profile.GetProperty("canClaimConformance").GetBoolean() == false);
            Assert.Contains(profiles.EnumerateArray(), profile =>
                string.Equals(profile.GetProperty("profile").GetString(), nameof(PdfComplianceProfile.PdfA3B), StringComparison.Ordinal) &&
                profile.GetProperty("requiredExternalValidators").EnumerateArray().Any(validator => string.Equals(validator.GetString(), nameof(PdfExternalValidatorKind.VeraPdf), StringComparison.Ordinal)) &&
                profile.GetProperty("externalValidatorProofs").EnumerateArray().Any(validator => string.Equals(validator.GetProperty("validatorKind").GetString(), nameof(PdfExternalValidatorKind.VeraPdf), StringComparison.Ordinal) &&
                    string.Equals(validator.GetProperty("status").GetString(), nameof(PdfExternalValidatorProofStatus.Missing), StringComparison.Ordinal) &&
                    validator.GetProperty("blocksConformanceClaim").GetBoolean()) &&
                profile.GetProperty("canClaimConformance").GetBoolean() == false);
            Assert.Contains(profiles.EnumerateArray(), profile =>
                string.Equals(profile.GetProperty("profile").GetString(), nameof(PdfComplianceProfile.Zugferd), StringComparison.Ordinal) &&
                profile.GetProperty("requiredExternalValidators").EnumerateArray().Any(validator => string.Equals(validator.GetString(), nameof(PdfExternalValidatorKind.Mustang), StringComparison.Ordinal)) &&
                profile.GetProperty("externalValidatorProofs").EnumerateArray().Any(validator => string.Equals(validator.GetProperty("validatorKind").GetString(), nameof(PdfExternalValidatorKind.Mustang), StringComparison.Ordinal) &&
                    string.Equals(validator.GetProperty("status").GetString(), nameof(PdfExternalValidatorProofStatus.Missing), StringComparison.Ordinal) &&
                    validator.GetProperty("blocksConformanceClaim").GetBoolean()) &&
                profile.GetProperty("canClaimConformance").GetBoolean() == false);
        } finally {
            Environment.SetEnvironmentVariable(ProofOutputEnv, string.IsNullOrEmpty(previousOutput) ? null : previousOutput);
            TryDeleteDirectory(directory);
        }
    }

    private static byte[] CreatePdfA3BFixture() {
        return CreatePdfA3BDocument().ToBytes();
    }

    private static PdfDocument CreatePdfA3BDocument() {
        return PdfDocument.Create(CreatePdfA3BOptions())
            .Meta(title: "OfficeIMO PDF/A-3b", author: "OfficeIMO")
            .Language("en-US")
            .SrgbOutputIntent()
            .AttachFile("source-data.xml", Encoding.UTF8.GetBytes("<source><id>123</id></source>"), "application/xml", PdfAssociatedFileRelationship.Data, "Source data")
            .H1("Archival PDF with Associated Data")
            .Paragraph(p => p.Text("This exact artifact is validated as PDF/A-3b before OfficeIMO claims conformance."));
    }

    private static byte[] CreatePdfA2BFixture() {
        return CreatePdfA2BDocument().ToBytes();
    }

    private static PdfDocument CreatePdfA2BDocument() {
        return PdfDocument.Create(CreatePdfA2BOptions())
            .Meta(title: "OfficeIMO PDF/A-2b", author: "OfficeIMO")
            .Language("en-US")
            .H1("Archival PDF")
            .Paragraph(p => p.Text("This exact artifact is validated as PDF/A-2b before OfficeIMO claims conformance."));
    }

    private static PdfOptions CreatePdfA3BOptions() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        return new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B, "en-US")
            .RequireCompliance(PdfComplianceProfile.PdfA3B)
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, fontData, "OfficeIMO Source Serif");
    }

    private static PdfOptions CreatePdfA2BOptions() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        return new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA2B, "en-US")
            .RequireCompliance(PdfComplianceProfile.PdfA2B)
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, fontData, "OfficeIMO Source Serif");
    }

    private static byte[] CreateElectronicInvoiceFixture(PdfComplianceProfile profile) {
        return CreateElectronicInvoiceDocument(profile).ToBytes();
    }

    private static PdfDocument CreateElectronicInvoiceDocument(PdfComplianceProfile profile) {
        byte[] invoiceXml = CreateElectronicInvoiceXml();
        return PdfDocument.Create(CreateEinvoiceOptions(profile, invoiceXml))
            .Meta(title: "OfficeIMO " + (profile == PdfComplianceProfile.FacturX ? "Factur-X" : "ZUGFeRD") + " Invoice", author: "OfficeIMO")
            .Language("en-US")
            .H1(profile == PdfComplianceProfile.FacturX ? "Factur-X Invoice" : "ZUGFeRD Invoice")
            .Paragraph(p => p.Text("The attached EN 16931 invoice XML and its PDF/A-3 carrier are validated as one exact artifact."));
    }

    private static PdfOptions CreateEinvoiceOptions(PdfComplianceProfile profile, byte[] invoiceXml) {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        return new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigureElectronicInvoiceGroundwork(
                profile,
                invoiceXml,
                relationship: PdfAssociatedFileRelationship.Alternative)
            .RequireCompliance(profile)
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, fontData, "OfficeIMO Source Serif");
    }

    private static byte[] CreateElectronicInvoiceXml() {
        return Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<rsm:CrossIndustryInvoice xmlns:rsm=\"urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100\" xmlns:ram=\"urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100\" xmlns:udt=\"urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100\">" +
            "<rsm:ExchangedDocumentContext>" +
            "<ram:GuidelineSpecifiedDocumentContextParameter>" +
            "<ram:ID>urn:cen.eu:en16931:2017</ram:ID>" +
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
            "<ram:GrossPriceProductTradePrice>" +
            "<ram:ChargeAmount>100.00</ram:ChargeAmount>" +
            "<ram:BasisQuantity unitCode=\"C62\">1</ram:BasisQuantity>" +
            "</ram:GrossPriceProductTradePrice>" +
            "<ram:NetPriceProductTradePrice>" +
            "<ram:ChargeAmount>100.00</ram:ChargeAmount>" +
            "<ram:BasisQuantity unitCode=\"C62\">1</ram:BasisQuantity>" +
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
            "<ram:PostalTradeAddress><ram:PostcodeCode>00-001</ram:PostcodeCode><ram:LineOne>Engine Street 1</ram:LineOne><ram:CityName>Warsaw</ram:CityName><ram:CountryID>PL</ram:CountryID></ram:PostalTradeAddress>" +
            "<ram:URIUniversalCommunication><ram:URIID schemeID=\"0088\">5901234123457</ram:URIID></ram:URIUniversalCommunication>" +
            "<ram:SpecifiedTaxRegistration><ram:ID schemeID=\"FC\">PL1234567890</ram:ID></ram:SpecifiedTaxRegistration>" +
            "<ram:SpecifiedTaxRegistration><ram:ID schemeID=\"VA\">PL1234567890</ram:ID></ram:SpecifiedTaxRegistration>" +
            "</ram:SellerTradeParty>" +
            "<ram:BuyerTradeParty>" +
            "<ram:Name>OfficeIMO Buyer</ram:Name>" +
            "<ram:PostalTradeAddress><ram:PostcodeCode>10115</ram:PostcodeCode><ram:LineOne>Archive Road 2</ram:LineOne><ram:CityName>Berlin</ram:CityName><ram:CountryID>DE</ram:CountryID></ram:PostalTradeAddress>" +
            "<ram:URIUniversalCommunication><ram:URIID schemeID=\"0088\">4006381333931</ram:URIID></ram:URIUniversalCommunication>" +
            "<ram:SpecifiedTaxRegistration><ram:ID schemeID=\"VA\">DE123456789</ram:ID></ram:SpecifiedTaxRegistration>" +
            "</ram:BuyerTradeParty>" +
            "</ram:ApplicableHeaderTradeAgreement>" +
            "<ram:ApplicableHeaderTradeDelivery>" +
            "<ram:ActualDeliverySupplyChainEvent><ram:OccurrenceDateTime><udt:DateTimeString format=\"102\">20260603</udt:DateTimeString></ram:OccurrenceDateTime></ram:ActualDeliverySupplyChainEvent>" +
            "</ram:ApplicableHeaderTradeDelivery>" +
            "<ram:ApplicableHeaderTradeSettlement>" +
            "<ram:PaymentReference>INV-2026-0001</ram:PaymentReference>" +
            "<ram:InvoiceCurrencyCode>EUR</ram:InvoiceCurrencyCode>" +
            "<ram:SpecifiedTradeSettlementPaymentMeans>" +
            "<ram:TypeCode>58</ram:TypeCode>" +
            "<ram:PayeePartyCreditorFinancialAccount>" +
            "<ram:IBANID>PL61109010140000071219812874</ram:IBANID>" +
            "</ram:PayeePartyCreditorFinancialAccount>" +
            "</ram:SpecifiedTradeSettlementPaymentMeans>" +
            "<ram:ApplicableTradeTax>" +
            "<ram:CalculatedAmount>23.00</ram:CalculatedAmount>" +
            "<ram:TypeCode>VAT</ram:TypeCode>" +
            "<ram:BasisAmount>100.00</ram:BasisAmount>" +
            "<ram:CategoryCode>S</ram:CategoryCode>" +
            "<ram:RateApplicablePercent>23</ram:RateApplicablePercent>" +
            "</ram:ApplicableTradeTax>" +
            "<ram:SpecifiedTradePaymentTerms>" +
            "<ram:Description>Due within 30 days</ram:Description>" +
            "<ram:DueDateDateTime><udt:DateTimeString format=\"102\">20260703</udt:DateTimeString></ram:DueDateDateTime>" +
            "</ram:SpecifiedTradePaymentTerms>" +
            "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
            "<ram:LineTotalAmount>100.00</ram:LineTotalAmount>" +
            "<ram:ChargeTotalAmount>0.00</ram:ChargeTotalAmount>" +
            "<ram:AllowanceTotalAmount>0.00</ram:AllowanceTotalAmount>" +
            "<ram:TaxBasisTotalAmount>100.00</ram:TaxBasisTotalAmount>" +
            "<ram:TaxTotalAmount currencyID=\"EUR\">23.00</ram:TaxTotalAmount>" +
            "<ram:GrandTotalAmount>123.00</ram:GrandTotalAmount>" +
            "<ram:DuePayableAmount>123.00</ram:DuePayableAmount>" +
            "</ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
            "</ram:ApplicableHeaderTradeSettlement>" +
            "</rsm:SupplyChainTradeTransaction>" +
            "</rsm:CrossIndustryInvoice>");
    }

    private static byte[] CreatePdfUa1Fixture() {
        return CreatePdfUa1Document().ToBytes();
    }

    private static PdfDocument CreatePdfUa1Document() {
        var accessibleFieldStyle = new PdfFormFieldStyle {
            AlternateName = "Contact email address"
        };
        var accessibleCheckBoxStyle = new PdfFormFieldStyle {
            AlternateName = "Receive compliance updates"
        };

        return PdfDocument.Create(CreatePdfUa1Options())
            .Meta(title: "OfficeIMO PDF/UA-1", author: "OfficeIMO")
            .H1("Accessible PDF")
            .Paragraph(p => p.Text("This fixture exercises reading order, Unicode text, accessible links, annotations, forms, and figures. ")
                .Link("Visit the OfficeIMO project", "https://officeimo.net/", contents: "OfficeIMO project website"))
            .Image(CreateOnePixelPng(), 12, 12, alternativeText: "Decorative sample pixel")
            .TextAnnotation("Accessibility review note", width: 18, height: 18)
            .TextField("Contact.Email", value: "reader@example.com", width: 180, height: 24, style: accessibleFieldStyle)
            .CheckBox("Contact.Updates", isChecked: true, size: 16, style: accessibleCheckBoxStyle);
    }

    private static byte[] CreateOnePixelPng() {
        return PdfPngTestImages.CreateRgbPng(1, 1);
    }

    private static void WriteProfileProofContract() {
        string? outputDirectory = GetProofOutputDirectory();
        if (outputDirectory == null) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        PdfComplianceArtifact pdfA2BArtifact = CreatePdfA2BDocument()
            .CreateComplianceArtifact(PdfComplianceProfile.PdfA2B);
        PdfComplianceArtifact pdfA3BArtifact = CreatePdfA3BDocument()
            .CreateComplianceArtifact(PdfComplianceProfile.PdfA3B);
        PdfComplianceArtifact facturXArtifact = CreateElectronicInvoiceDocument(PdfComplianceProfile.FacturX)
            .CreateComplianceArtifact(PdfComplianceProfile.FacturX);
        PdfComplianceArtifact zugferdArtifact = CreateElectronicInvoiceDocument(PdfComplianceProfile.Zugferd)
            .CreateComplianceArtifact(PdfComplianceProfile.Zugferd);
        PdfComplianceArtifact pdfUa1Artifact = CreatePdfUa1Document()
            .CreateComplianceArtifact(PdfComplianceProfile.PdfUa1);
        var contract = new ProfileProofContract {
            SchemaVersion = 4,
            GeneratedBy = nameof(PdfComplianceArtifact) + "." + nameof(PdfComplianceArtifact.AssessProof),
            ExternalEvidenceMode = "ExactArtifactValidationInjectedByProofExporter",
            Profiles = new[] {
                CreateProofContractRow(pdfA2BArtifact),
                CreateProofContractRow(pdfA3BArtifact),
                CreateProofContractRow(pdfUa1Artifact),
                CreateProofContractRow(facturXArtifact),
                CreateProofContractRow(zugferdArtifact)
            }
        };
        string json = JsonSerializer.Serialize(contract, new JsonSerializerOptions {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        });
        File.WriteAllText(Path.Combine(outputDirectory, "officeimo-profile-proof-contract.json"), json, Encoding.UTF8);
    }

    private static ProfileProofContractRow CreateProofContractRow(PdfComplianceArtifact artifact) =>
        CreateProofContractRow(artifact.AssessProof(Array.Empty<PdfExternalValidationResult>()));

    private static ProfileProofContractRow CreateProofContractRow(PdfComplianceProofReport proof) {
        return new ProfileProofContractRow {
            Profile = proof.Profile.ToString(),
            DisplayName = proof.DisplayName,
            IsInternallyReady = proof.IsInternallyReady,
            HasRequiredExternalValidation = proof.HasRequiredExternalValidation,
            CanClaimConformance = proof.CanClaimConformance,
            ProofStatus = proof.ProofStatus,
            HasArtifactEvidence = proof.HasArtifactEvidence,
            ArtifactSha256 = proof.ArtifactSha256,
            ArtifactSizeBytes = proof.ArtifactSizeBytes,
            RequiredExternalValidators = proof.RequiredExternalValidators.Select(validator => validator.ToString()).ToArray(),
            MissingExternalValidators = proof.MissingExternalValidators.Select(validator => validator.ToString()).ToArray(),
            FailedExternalValidationCount = proof.FailedExternalValidations.Count,
            ExternalValidatorProofs = proof.ExternalValidatorProofs.Select(row => new ExternalValidatorProofContractRow {
                ValidatorKind = row.ValidatorKind.ToString(),
                Status = row.Status.ToString(),
                IsSatisfied = row.IsSatisfied,
                BlocksConformanceClaim = row.BlocksConformanceClaim,
                ValidatorName = row.ValidatorName,
                Diagnostic = row.Diagnostic,
                Profile = row.Profile,
                ExitCode = row.ExitCode,
                ValidatorVersion = row.ValidatorVersion,
                ArtifactSha256 = row.ArtifactSha256,
                ArtifactSizeBytes = row.ArtifactSizeBytes,
                ValidatedAtUtc = row.ValidatedAtUtc,
                Warnings = row.Warnings.ToArray()
            }).ToArray(),
            MissingRequirementIds = proof.MissingRequirements.Select(requirement => requirement.Id).ToArray(),
            UnsupportedRequirementIds = proof.UnsupportedRequirements.Select(requirement => requirement.Id).ToArray()
        };
    }

    private static PdfOptions CreatePdfUa1Options() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        return new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigurePdfUaGroundwork("en-US")
            .RequireCompliance(PdfComplianceProfile.PdfUa1)
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, fontData, "OfficeIMO Source Serif");
    }

    private static void WriteProofPdf(string fileName, byte[] bytes) {
        string? outputDirectory = GetProofOutputDirectory();
        if (outputDirectory == null) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        File.WriteAllBytes(Path.Combine(outputDirectory, fileName), bytes);
    }

    private static void WriteProofText(string fileName, string text) {
        string? outputDirectory = GetProofOutputDirectory();
        if (outputDirectory == null) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        File.WriteAllText(Path.Combine(outputDirectory, fileName), text, Encoding.UTF8);
    }

    private static string? GetProofOutputDirectory() {
        string? outputDirectory = Environment.GetEnvironmentVariable(ProofOutputEnv);
        return string.IsNullOrWhiteSpace(outputDirectory) ? null : outputDirectory;
    }

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) {
                Directory.Delete(path, recursive: true);
            }
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }
    }

    private sealed class ProfileProofContract {
        public int SchemaVersion { get; set; }

        public string GeneratedBy { get; set; } = string.Empty;

        public string ExternalEvidenceMode { get; set; } = string.Empty;

        public ProfileProofContractRow[] Profiles { get; set; } = Array.Empty<ProfileProofContractRow>();
    }

    private sealed class ProfileProofContractRow {
        public string Profile { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public bool IsInternallyReady { get; set; }

        public bool HasRequiredExternalValidation { get; set; }

        public bool CanClaimConformance { get; set; }

        public string ProofStatus { get; set; } = string.Empty;

        public bool HasArtifactEvidence { get; set; }

        public string? ArtifactSha256 { get; set; }

        public long? ArtifactSizeBytes { get; set; }

        public string[] RequiredExternalValidators { get; set; } = Array.Empty<string>();

        public string[] MissingExternalValidators { get; set; } = Array.Empty<string>();

        public int FailedExternalValidationCount { get; set; }

        public ExternalValidatorProofContractRow[] ExternalValidatorProofs { get; set; } = Array.Empty<ExternalValidatorProofContractRow>();

        public string[] MissingRequirementIds { get; set; } = Array.Empty<string>();

        public string[] UnsupportedRequirementIds { get; set; } = Array.Empty<string>();
    }

    private sealed class ExternalValidatorProofContractRow {
        public string ValidatorKind { get; set; } = string.Empty;

        public string Status { get; set; } = string.Empty;

        public bool IsSatisfied { get; set; }

        public bool BlocksConformanceClaim { get; set; }

        public string ValidatorName { get; set; } = string.Empty;

        public string Diagnostic { get; set; } = string.Empty;

        public string? Profile { get; set; }

        public int? ExitCode { get; set; }

        public string? ValidatorVersion { get; set; }

        public string? ArtifactSha256 { get; set; }

        public long? ArtifactSizeBytes { get; set; }

        public DateTimeOffset? ValidatedAtUtc { get; set; }

        public string[] Warnings { get; set; } = Array.Empty<string>();
    }
}
