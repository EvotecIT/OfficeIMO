using System;
using System.IO;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComplianceGateTests {
    private const string ProofOutputEnv = "OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT";
    private static readonly PdfStandardFont[] GeneratedGroundworkFonts = {
        PdfStandardFont.Helvetica,
        PdfStandardFont.HelveticaBold
    };

    [Fact]
    public void PdfA3GroundworkFixture_ContainsCurrentArchivalPrimitivesWithoutEnablingFormalProfile() {
        byte[] bytes = CreatePdfA3GroundworkFixture();
        WriteProofPdf("officeimo-pdfa3-groundwork.pdf", bytes);
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
    }

    [Fact]
    public void VeraPdfGate_RejectsGroundworkFixtureUntilFormalPdfA3BGenerationExists() {
        byte[] fixture = CreatePdfA3GroundworkFixture();
        WriteProofPdf("officeimo-pdfa3-groundwork.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf();
        if (!validator.IsAvailable) {
            WriteProofText("verapdf-pdfa3-groundwork.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-pdfa3-groundwork.pdf");
        WriteProofText("verapdf-pdfa3-groundwork.txt", result.GetDiagnosticText());

        Assert.False(result.ExitCode == 0, "The PDF/A-3b groundwork fixture unexpectedly passed veraPDF. Enable the formal OfficeIMO.Pdf profile only after the compliance profile itself is implemented and the validator fixture is intentionally flipped to expect success." + Environment.NewLine + result.GetDiagnosticText());
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void MustangGate_RejectsEinvoiceGroundworkFixtureUntilFacturXProfileGenerationExists() {
        byte[] fixture = CreateEinvoiceGroundworkFixture();
        WriteProofPdf("officeimo-einvoice-groundwork.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.Mustang();
        if (!validator.IsAvailable) {
            WriteProofText("mustang-einvoice-groundwork.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-einvoice-groundwork.pdf");
        WriteProofText("mustang-einvoice-groundwork.txt", result.GetDiagnosticText());

        Assert.False(result.ExitCode == 0, "The e-invoice groundwork fixture unexpectedly passed Mustang validation. Enable Factur-X/ZUGFeRD profile output only after profile-specific XML, XMP, PDF/A-3, and validator evidence are intentionally implemented." + Environment.NewLine + result.GetDiagnosticText());
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void VeraPdfGate_RejectsEinvoiceGroundworkFixtureUntilFacturXProfileGenerationExists() {
        byte[] fixture = CreateEinvoiceGroundworkFixture();
        WriteProofPdf("officeimo-einvoice-groundwork.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf();
        if (!validator.IsAvailable) {
            WriteProofText("verapdf-einvoice-groundwork.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-einvoice-groundwork.pdf");
        WriteProofText("verapdf-einvoice-groundwork.txt", result.GetDiagnosticText());

        Assert.False(result.ExitCode == 0, "The e-invoice groundwork fixture unexpectedly passed veraPDF. Enable Factur-X/ZUGFeRD profile output only after the actual e-invoice PDF/A-3 carrier is intentionally implemented and validated." + Environment.NewLine + result.GetDiagnosticText());
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void PdfUaGroundworkFixture_ContainsCurrentAccessibilityPrimitivesWithoutEnablingFormalProfile() {
        byte[] bytes = CreatePdfUaGroundworkFixture();
        WriteProofPdf("officeimo-pdfua-groundwork.pdf", bytes);
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
    }

    [Fact]
    public void PdfUaValidatorGate_RejectsGroundworkFixtureUntilFormalPdfUaGenerationExists() {
        byte[] fixture = CreatePdfUaGroundworkFixture();
        WriteProofPdf("officeimo-pdfua-groundwork.pdf", fixture);
        PdfExternalValidator validator = PdfExternalValidator.PdfUa();
        if (!validator.IsAvailable) {
            WriteProofText("pdfua-groundwork.txt", validator.GetNotConfiguredText());
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalProcessResult result = validator.Run(fixture, "officeimo-pdfua-groundwork.pdf");
        WriteProofText("pdfua-groundwork.txt", result.GetDiagnosticText());

        Assert.False(result.ExitCode == 0, "The PDF/UA groundwork fixture unexpectedly passed the configured PDF/UA validator. Enable the formal OfficeIMO.Pdf profile only after tagged structure, reading order, alternate text, font mapping, and validator evidence are intentionally implemented." + Environment.NewLine + result.GetDiagnosticText());
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA3B, "veraPDF validation fixtures in the build lane")]
    [InlineData(PdfComplianceProfile.PdfUa1, "PDF/UA identification XMP")]
    [InlineData(PdfComplianceProfile.FacturX, "Mustang validation fixtures in the build lane")]
    [InlineData(PdfComplianceProfile.Zugferd, "Mustang validation fixtures in the build lane")]
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
    public void ProofOutputHook_WritesGroundworkFixturesWhenConfigured() {
        string previousOutput = Environment.GetEnvironmentVariable(ProofOutputEnv) ?? string.Empty;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfComplianceProof", Guid.NewGuid().ToString("N"));
        try {
            Environment.SetEnvironmentVariable(ProofOutputEnv, directory);

            WriteProofPdf("officeimo-pdfa3-groundwork.pdf", CreatePdfA3GroundworkFixture());
            WriteProofPdf("officeimo-einvoice-groundwork.pdf", CreateEinvoiceGroundworkFixture());
            WriteProofPdf("officeimo-pdfua-groundwork.pdf", CreatePdfUaGroundworkFixture());
            WriteProfileProofContract();
            WriteProofText("verapdf-pdfa3-groundwork.txt", "validator diagnostic");
            WriteProofText("verapdf-einvoice-groundwork.txt", "validator diagnostic");
            WriteProofText("pdfua-groundwork.txt", "validator diagnostic");

            Assert.True(File.Exists(Path.Combine(directory, "officeimo-pdfa3-groundwork.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-einvoice-groundwork.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-pdfua-groundwork.pdf")));
            Assert.True(File.Exists(Path.Combine(directory, "officeimo-profile-proof-contract.json")));
            Assert.True(File.Exists(Path.Combine(directory, "verapdf-pdfa3-groundwork.txt")));
            Assert.True(File.Exists(Path.Combine(directory, "verapdf-einvoice-groundwork.txt")));
            Assert.True(File.Exists(Path.Combine(directory, "pdfua-groundwork.txt")));
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-pdfa3-groundwork.pdf")).Length > 0);
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-einvoice-groundwork.pdf")).Length > 0);
            Assert.True(new FileInfo(Path.Combine(directory, "officeimo-pdfua-groundwork.pdf")).Length > 0);
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
            Assert.Equal(2, root.GetProperty("schemaVersion").GetInt32());
            Assert.Equal("NoExternalValidationInjected", root.GetProperty("externalEvidenceMode").GetString());
            JsonElement profiles = root.GetProperty("profiles");
            Assert.Equal(4, profiles.GetArrayLength());
            Assert.Contains(profiles.EnumerateArray(), profile =>
                string.Equals(profile.GetProperty("profile").GetString(), nameof(PdfComplianceProfile.PdfA3B), StringComparison.Ordinal) &&
                profile.GetProperty("requiredExternalValidators").EnumerateArray().Any(validator => string.Equals(validator.GetString(), nameof(PdfExternalValidatorKind.VeraPdf), StringComparison.Ordinal)) &&
                profile.GetProperty("canClaimConformance").GetBoolean() == false);
            Assert.Contains(profiles.EnumerateArray(), profile =>
                string.Equals(profile.GetProperty("profile").GetString(), nameof(PdfComplianceProfile.Zugferd), StringComparison.Ordinal) &&
                profile.GetProperty("requiredExternalValidators").EnumerateArray().Any(validator => string.Equals(validator.GetString(), nameof(PdfExternalValidatorKind.Mustang), StringComparison.Ordinal)) &&
                profile.GetProperty("canClaimConformance").GetBoolean() == false);
        } finally {
            Environment.SetEnvironmentVariable(ProofOutputEnv, string.IsNullOrEmpty(previousOutput) ? null : previousOutput);
            TryDeleteDirectory(directory);
        }
    }

    private static byte[] CreatePdfA3GroundworkFixture() {
        return PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .PdfAIdentification(3, "B")
            .Meta(title: "OfficeIMO Archival Groundwork", author: "OfficeIMO")
            .Language("en-US")
            .SrgbOutputIntent()
            .AttachFile("source-data.xml", Encoding.UTF8.GetBytes("<source><id>123</id></source>"), "application/xml", PdfAssociatedFileRelationship.Data, "Source data")
            .H1("Archival Groundwork")
            .Paragraph(p => p.Text("This fixture intentionally contains compliance primitives but does not claim PDF/A conformance."))
            .ToBytes();
    }

    private static byte[] CreateEinvoiceGroundworkFixture() {
        byte[] invoiceXml = CreateEinvoiceGroundworkXml();
        return PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .PdfAIdentification(3, "B")
            .ElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .Meta(title: "OfficeIMO E-Invoice Groundwork", author: "OfficeIMO")
            .Language("en-US")
            .SrgbOutputIntent()
            .AttachFile("factur-x.xml", invoiceXml, "application/xml", PdfAssociatedFileRelationship.Data, "EN 16931 XML payload placeholder")
            .H1("E-Invoice Groundwork")
            .Paragraph(p => p.Text("This fixture intentionally exercises associated-file output without claiming Factur-X or ZUGFeRD conformance."))
            .ToBytes();
    }

    private static byte[] CreateEinvoiceGroundworkXml() {
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

    private static byte[] CreatePdfUaGroundworkFixture() {
        return PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .ConfigurePdfUaGroundwork("en-US")
            .Meta(title: "OfficeIMO PDF/UA Groundwork", author: "OfficeIMO")
            .H1("Accessibility Groundwork")
            .Paragraph(p => p.Text("This fixture intentionally contains PDF/UA primitives but does not claim PDF/UA conformance."))
            .Image(CreateOnePixelPng(), 12, 12, alternativeText: "Decorative sample pixel")
            .ToBytes();
    }

    private static byte[] CreateOnePixelPng() {
        return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
    }

    private static void WriteProfileProofContract() {
        string? outputDirectory = GetProofOutputDirectory();
        if (outputDirectory == null) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        var contract = new ProfileProofContract {
            SchemaVersion = 2,
            GeneratedBy = nameof(PdfComplianceAnalyzer) + "." + nameof(PdfComplianceAnalyzer.AssessProof),
            ExternalEvidenceMode = "NoExternalValidationInjected",
            Profiles = new[] {
                CreateProofContractRow(PdfComplianceProfile.PdfA3B, CreatePdfA3GroundworkOptions(), GeneratedGroundworkFonts),
                CreateProofContractRow(PdfComplianceProfile.PdfUa1, CreatePdfUaGroundworkOptions(), GeneratedGroundworkFonts),
                CreateProofContractRow(PdfComplianceProfile.FacturX, CreateEinvoiceGroundworkOptions(), GeneratedGroundworkFonts),
                CreateProofContractRow(PdfComplianceProfile.Zugferd, CreateEinvoiceGroundworkOptions(), GeneratedGroundworkFonts)
            }
        };
        string json = JsonSerializer.Serialize(contract, new JsonSerializerOptions {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        });
        File.WriteAllText(Path.Combine(outputDirectory, "officeimo-profile-proof-contract.json"), json, Encoding.UTF8);
    }

    private static ProfileProofContractRow CreateProofContractRow(PdfComplianceProfile profile, PdfOptions options, IEnumerable<PdfStandardFont> generatedStandardFonts) {
        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            profile,
            options,
            externalValidations: Array.Empty<PdfExternalValidationResult>(),
            generatedStandardFonts: generatedStandardFonts);

        return new ProfileProofContractRow {
            Profile = proof.Profile.ToString(),
            DisplayName = proof.DisplayName,
            IsInternallyReady = proof.IsInternallyReady,
            HasRequiredExternalValidation = proof.HasRequiredExternalValidation,
            CanClaimConformance = proof.CanClaimConformance,
            RequiredExternalValidators = proof.RequiredExternalValidators.Select(validator => validator.ToString()).ToArray(),
            MissingExternalValidators = proof.MissingExternalValidators.Select(validator => validator.ToString()).ToArray(),
            FailedExternalValidationCount = proof.FailedExternalValidations.Count,
            MissingRequirementIds = proof.Readiness.MissingRequirements.Select(requirement => requirement.Id).ToArray(),
            UnsupportedRequirementIds = proof.Readiness.UnsupportedRequirements.Select(requirement => requirement.Id).ToArray()
        };
    }

    private static PdfOptions CreatePdfA3GroundworkOptions() {
        return new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            IncludeStandardFontToUnicodeMaps = true
        }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();
    }

    private static PdfOptions CreatePdfUaGroundworkOptions() {
        return new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigurePdfUaGroundwork("en-US");
    }

    private static PdfOptions CreateEinvoiceGroundworkOptions() {
        return new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            IncludeStandardFontToUnicodeMaps = true
        }
            .SetPdfAIdentification(3, "B")
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", CreateEinvoiceGroundworkXml(), "application/xml", PdfAssociatedFileRelationship.Data, "EN 16931 XML payload placeholder");
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

        public string[] RequiredExternalValidators { get; set; } = Array.Empty<string>();

        public string[] MissingExternalValidators { get; set; } = Array.Empty<string>();

        public int FailedExternalValidationCount { get; set; }

        public string[] MissingRequirementIds { get; set; } = Array.Empty<string>();

        public string[] UnsupportedRequirementIds { get; set; } = Array.Empty<string>();
    }
}
