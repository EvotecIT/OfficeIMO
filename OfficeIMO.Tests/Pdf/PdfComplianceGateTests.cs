using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComplianceGateTests {
    [Fact]
    public void PdfA3GroundworkFixture_ContainsCurrentArchivalPrimitivesWithoutEnablingFormalProfile() {
        byte[] bytes = CreatePdfA3GroundworkFixture();
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
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf();
        if (!validator.IsAvailable) {
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalValidationResult result = validator.Run(CreatePdfA3GroundworkFixture(), "officeimo-pdfa3-groundwork.pdf");

        Assert.False(result.ExitCode == 0, "The PDF/A-3b groundwork fixture unexpectedly passed veraPDF. Enable the formal OfficeIMO.Pdf profile only after the compliance profile itself is implemented and the validator fixture is intentionally flipped to expect success." + Environment.NewLine + result.GetDiagnosticText());
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Fact]
    public void MustangGate_RejectsEinvoiceGroundworkFixtureUntilFacturXProfileGenerationExists() {
        PdfExternalValidator validator = PdfExternalValidator.Mustang();
        if (!validator.IsAvailable) {
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }

        PdfExternalValidationResult result = validator.Run(CreateEinvoiceGroundworkFixture(), "officeimo-einvoice-groundwork.pdf");

        Assert.False(result.ExitCode == 0, "The e-invoice groundwork fixture unexpectedly passed Mustang validation. Enable Factur-X/ZUGFeRD profile output only after profile-specific XML, XMP, PDF/A-3, and validator evidence are intentionally implemented." + Environment.NewLine + result.GetDiagnosticText());
        Assert.NotEmpty(result.GetDiagnosticText());
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA3B, "veraPDF validation fixtures in the build lane")]
    [InlineData(PdfComplianceProfile.FacturX, "Mustang validation fixtures in the build lane")]
    [InlineData(PdfComplianceProfile.Zugferd, "Mustang validation fixtures in the build lane")]
    public void FormalProfiles_StillFailClosedUntilValidatorBackedGenerationExists(PdfComplianceProfile profile, string expectedRequirement) {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDoc.Create()
                .Compliance(profile)
                .Meta(title: "Compliance profile guard")
                .Paragraph(p => p.Text("Profile output must fail closed until validator evidence exists."))
                .ToBytes());

        Assert.Contains("cannot yet generate certified", exception.Message, StringComparison.Ordinal);
        Assert.Contains(expectedRequirement, exception.Message, StringComparison.Ordinal);
    }

    private static byte[] CreatePdfA3GroundworkFixture() {
        return PdfDoc.Create(new PdfOptions {
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
        byte[] invoiceXml = Encoding.UTF8.GetBytes(
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

        return PdfDoc.Create(new PdfOptions {
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

    private sealed class PdfExternalValidator {
        private const string RequireEnv = "OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS";
        private readonly string[] _arguments;

        private PdfExternalValidator(string name, string? executablePath, string[] arguments, bool autoDetected) {
            Name = name;
            ExecutablePath = executablePath;
            _arguments = arguments;
            AutoDetected = autoDetected;
        }

        internal string Name { get; }

        internal string? ExecutablePath { get; }

        internal bool AutoDetected { get; }

        internal bool IsAvailable => !string.IsNullOrWhiteSpace(ExecutablePath);

        internal static PdfExternalValidator VeraPdf() {
            string? explicitPath = FirstNonEmpty(
                Environment.GetEnvironmentVariable("OFFICEIMO_VERAPDF"),
                Environment.GetEnvironmentVariable("OFFICEIMO_VERAPDF_PATH"));
            string? path = explicitPath ?? FindOnPath("verapdf", "verapdf.bat", "verapdf.exe");
            string[] args = GetConfiguredArgs("OFFICEIMO_VERAPDF_ARGS", "-f", "3b", "{pdf}");
            return new PdfExternalValidator("veraPDF", path, args, explicitPath == null && path != null);
        }

        internal static PdfExternalValidator Mustang() {
            string? explicitPath = FirstNonEmpty(
                Environment.GetEnvironmentVariable("OFFICEIMO_MUSTANG"),
                Environment.GetEnvironmentVariable("OFFICEIMO_MUSTANG_PATH"));

            string? path = explicitPath ?? FindOnPath("mustangproject", "mustangproject.bat", "mustangproject.exe", "mustang", "mustang.bat", "mustang.exe");
            string[] args = GetConfiguredArgs("OFFICEIMO_MUSTANG_ARGS", "--action", "validate", "--source", "{pdf}");
            if (path != null && string.Equals(Path.GetExtension(path), ".jar", StringComparison.OrdinalIgnoreCase)) {
                args = new[] { "-jar", path }.Concat(args).ToArray();
                path = FindOnPath("java", "java.exe");
            }

            return new PdfExternalValidator("Mustang", path, args, explicitPath == null && path != null);
        }

        internal static void SkipUnlessRequired(PdfExternalValidator validator) {
            if (IsRequired()) {
                throw new InvalidOperationException(
                    validator.Name + " compliance validation was required, but the validator was not found. Set OFFICEIMO_VERAPDF, OFFICEIMO_VERAPDF_PATH, OFFICEIMO_MUSTANG, or OFFICEIMO_MUSTANG_PATH as appropriate, or add the tool to PATH.");
            }
        }

        internal PdfExternalValidationResult Run(byte[] pdfBytes, string fileName) {
            if (!IsAvailable) {
                throw new InvalidOperationException(Name + " validator is not configured.");
            }

            string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfCompliance", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(workDir);
            string pdfPath = Path.Combine(workDir, fileName);
            try {
                File.WriteAllBytes(pdfPath, pdfBytes);
                string arguments = string.Join(" ", _arguments.Select(argument => QuoteArgument(argument.Replace("{pdf}", pdfPath))));
                var startInfo = new ProcessStartInfo {
                    FileName = ExecutablePath!,
                    Arguments = arguments,
                    WorkingDirectory = workDir,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using Process process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start " + Name + " validator.");
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                if (!process.WaitForExit(60000)) {
                    try {
                        process.Kill();
                    } catch (InvalidOperationException) {
                    }

                    throw new TimeoutException(Name + " validation did not finish within 60 seconds.");
                }

                return new PdfExternalValidationResult(Name, ExecutablePath!, arguments, process.ExitCode, output, error, AutoDetected);
            } finally {
                TryDeleteDirectory(workDir);
            }
        }

        private static bool IsRequired() =>
            string.Equals(Environment.GetEnvironmentVariable(RequireEnv), "1", StringComparison.Ordinal);

        private static string[] GetConfiguredArgs(string envName, params string[] defaultArgs) {
            string? raw = Environment.GetEnvironmentVariable(envName);
            return string.IsNullOrWhiteSpace(raw)
                ? defaultArgs
                : SplitCommandLine(raw!);
        }

        private static string? FirstNonEmpty(params string?[] values) {
            foreach (string? value in values) {
                if (!string.IsNullOrWhiteSpace(value)) {
                    return value;
                }
            }

            return null;
        }

        private static string? FindOnPath(params string[] names) {
            string? path = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrWhiteSpace(path)) {
                return null;
            }

            foreach (string directory in path.Split(Path.PathSeparator)) {
                if (string.IsNullOrWhiteSpace(directory)) {
                    continue;
                }

                foreach (string name in names) {
                    string candidate = Path.Combine(directory, name);
                    if (File.Exists(candidate)) {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private static string[] SplitCommandLine(string value) {
            var args = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            foreach (char c in value) {
                if (c == '"') {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (char.IsWhiteSpace(c) && !inQuotes) {
                    if (current.Length > 0) {
                        args.Add(current.ToString());
                        current.Clear();
                    }

                    continue;
                }

                current.Append(c);
            }

            if (current.Length > 0) {
                args.Add(current.ToString());
            }

            return args.ToArray();
        }

        private static string QuoteArgument(string argument) {
            if (argument.Length == 0) {
                return "\"\"";
            }

            if (argument.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) {
                return argument;
            }

            return "\"" + argument.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
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
    }

    private sealed class PdfExternalValidationResult {
        internal PdfExternalValidationResult(string validatorName, string executablePath, string arguments, int exitCode, string output, string error, bool autoDetected) {
            ValidatorName = validatorName;
            ExecutablePath = executablePath;
            Arguments = arguments;
            ExitCode = exitCode;
            Output = output;
            Error = error;
            AutoDetected = autoDetected;
        }

        internal string ValidatorName { get; }

        internal string ExecutablePath { get; }

        internal string Arguments { get; }

        internal int ExitCode { get; }

        internal string Output { get; }

        internal string Error { get; }

        internal bool AutoDetected { get; }

        internal string GetDiagnosticText() {
            var sb = new StringBuilder();
            sb.Append(ValidatorName)
                .Append(" exited with code ")
                .Append(ExitCode.ToString(CultureInfo.InvariantCulture))
                .Append(" using ")
                .Append(ExecutablePath)
                .Append(' ')
                .Append(Arguments)
                .AppendLine();
            if (!string.IsNullOrWhiteSpace(Output)) {
                sb.AppendLine(Output.Trim());
            }

            if (!string.IsNullOrWhiteSpace(Error)) {
                sb.AppendLine(Error.Trim());
            }

            return sb.ToString();
        }
    }
}
