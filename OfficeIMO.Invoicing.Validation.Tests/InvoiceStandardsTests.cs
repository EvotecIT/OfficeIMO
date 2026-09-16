using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public sealed class InvoiceStandardsTheoryAttribute : TheoryAttribute {
    public InvoiceStandardsTheoryAttribute() {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_STANDARDS_TESTS") != "1") Skip = "Opt-in authority artifact validation; run Build/Test-InvoicingStandards.ps1.";
    }
}
public partial class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task ExplicitlyLossyAccountingBreakdownRewritePassesAuthorityRules(InvoiceSyntax target) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 25m;
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl))));
        XElement[] totals = document.Root!.Elements().Where(e => e.Name.LocalName == "TaxTotal").ToArray();
        XElement subtotal = new XElement(totals[0].Elements().Single(e => e.Name.LocalName == "TaxSubtotal"));
        foreach (XAttribute currency in subtotal.Descendants().Attributes("currencyID")) currency.Value = "USD";
        totals[1].Add(subtotal);
        InvoiceReadResult read = InvoiceParser.Read(Encoding.UTF8.GetBytes(document.ToString()));
        Assert.False(read.HasCompleteMapping);
        byte[] rewritten = read.Write(InvoiceTestContracts.En16931(target), allowUnmappedDataLoss: true);
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(rewritten, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, "ftp://example.test/document.pdf")]
    [InlineData(InvoiceSyntax.Ubl, "ftp://example.test/document.pdf")]
    [InlineData(InvoiceSyntax.Cii, "urn:uuid:00112233-4455-6677-8899-aabbccddeeff")]
    [InlineData(InvoiceSyntax.Ubl, "urn:uuid:00112233-4455-6677-8899-aabbccddeeff")]
    public async Task NonHttpSupportingDocumentLocationsPassAuthorityRules(InvoiceSyntax syntax, string uri) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = "support", ExternalUri = uri });
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData(0)]
    [InlineData(1)]
    public async Task SingletonUblPaymentReferenceAndItsConversionPassAuthorityRules(int referenceIndex) {
        Invoice invoice = InvoiceFixture.Create();
        string reference = invoice.Payments[0].Reference!;
        invoice.Payments[0].Reference = referenceIndex == 0 ? reference : null;
        invoice.Payments.Add(new InvoicePayment { MeansCode = invoice.Payments[0].MeansCode, Reference = referenceIndex == 1 ? reference : null,
            Account = new InvoiceBankAccount { Identifier = "DE89370400440532013000" } });
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        var validator = new InvoiceValidator(Bundle(), Runner());
        InvoiceValidationReport original = await validator.ValidateAsync(xml, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(original.IsValid, Report(original));
        InvoiceConversionResult conversion = InvoiceConverter.Convert(xml, InvoiceTestContracts.En16931(InvoiceSyntax.Cii));
        Assert.True(conversion.Succeeded, string.Join("; ", conversion.Diagnostics.Select(d => d.Message)));
        InvoiceValidationReport converted = await validator.ValidateAsync(conversion.Xml!, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(converted.IsValid, Report(converted));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task MultiplePaymentOccurrencesWithAccountsPassAuthorityRules(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        InvoicePayment payment = invoice.Payments[0];
        invoice.Payments.Add(new InvoicePayment { MeansCode = payment.MeansCode,
            Account = new InvoiceBankAccount { Identifier = "DE89370400440532013000", Name = "Second account" } });
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ExcessiveContentFailuresRemainInvalidInsteadOfEngineFailure(bool schemaErrors) {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Create(), InvoiceTestContracts.En16931(InvoiceSyntax.Ubl))));
        XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
        XNamespace cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
        XElement template = document.Root!.Element(cac + "InvoiceLine")!;
        template.Remove();
        for (int index = 0; index < 1100; index++) {
            var line = new XElement(template);
            line.Element(cbc + "ID")!.Value = (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (schemaErrors) line.SetAttributeValue("invalid-attribute", "invalid");
            else line.Element(cac + "Item")!.Element(cbc + "Name")!.Remove();
            document.Root.Add(line);
        }
        byte[] xml = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.Equal(schemaErrors ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(schemaErrors ? InvoiceValidationStatus.NotRun : InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, d => d.Code == "INV-DIAGNOSTICS-TRUNCATED" && d.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.DoesNotContain(report.Diagnostics, d => d.Code == "INV-SCHEMA-ENGINE" || d.Code == "INV-RULES-ENGINE");
        Assert.InRange(report.Diagnostics.Count, 1, 1000);
    }

    private static InvoiceRuleBundle Bundle() => InvoiceRuleBundle.Load(
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE") ?? throw new InvalidOperationException("Rule bundle path is required."),
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PEPPOL_RULES"),
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_FACTURX_RULE_BUNDLE"));
    private static SaxonInvoiceRulesRunner Runner() => new SaxonInvoiceRulesRunner(
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR") ?? throw new InvalidOperationException("Saxon path is required."),
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java");

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSpecificationRelease.En16931_1_3_16)]
    public async Task MissingRuntimeIsAnEngineFailureAndCancellationIsPreserved(InvoiceSpecificationRelease release) {
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create(), InvoiceTestContracts.En16931());
        string jar = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!;
        var validator = new InvoiceValidator(Bundle(), new SaxonInvoiceRulesRunner(jar, "officeimo-nonexistent-java-" + Guid.NewGuid().ToString("N")));
        InvoiceValidationReport report = await validator.ValidateAsync(xml, release);
        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Failed, report.BusinessRulesStatus);
        Assert.Null(report.Runner);
        Assert.False(report.IsValid);
        Assert.Contains(report.Diagnostics, d => d.Code == "INV-RULES-ENGINE");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => validator.ValidateAsync(xml, release, cancelled.Token));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.En16931, InvoiceSpecificationRelease.En16931_1_3_16)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.En16931, InvoiceSpecificationRelease.En16931_1_3_16)]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, InvoiceSpecificationRelease.PeppolBis_3_0_21)]
    public async Task GeneratedInvoicesPassPinnedSchemaAndBusinessRules(InvoiceSyntax syntax, InvoiceProfile profile, InvoiceSpecificationRelease release) {
        var invoice = InvoiceFixture.Create();
        if (profile == InvoiceProfile.PeppolBis) {
            invoice.Seller.ElectronicAddress = new InvoiceIdentifier("1234567890128", "0088");
            invoice.Buyer.ElectronicAddress = new InvoiceIdentifier("1234567890135", "0088");
        }
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.For(syntax, profile));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, release);
        Assert.True(report.IsValid, Report(report));
        Assert.Equal(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(xml)), report.Sha256);
        Assert.Equal(Runner().Identity, report.Runner);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_EVIDENCE");
        if (output != null) {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, syntax + "-" + profile + ".xml"), xml);
            await File.WriteAllTextAsync(Path.Combine(output, syntax + "-" + profile + ".json"), System.Text.Json.JsonSerializer.Serialize(report, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
        }
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public async Task GeneratedFacturXProfilesPassPinnedSchemaAndBusinessRules(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.BusinessProcessId = null;
        var options = InvoiceTestContracts.FacturX(profile,
            profile == InvoiceProfile.En16931 || profile == InvoiceProfile.Extended
                ? InvoiceProjectionPolicy.RejectDataLoss
                : InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        var validator = new InvoiceValidator(Bundle(), Runner());
        InvoiceValidationReport report = await validator.ValidateAsync(xml, InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2);
        Assert.True(report.IsValid, Report(report));
        Assert.Equal(InvoiceRuleBundle.FacturXSha256, report.AuthorityBundleSha256);
        Assert.Equal(InvoiceRuleBundle.FacturXSourceCommit, report.AuthoritySourceCommit);
        InvoiceProfileDeclaration declaration = InvoiceProfileDeclaration.Read(xml);
        Assert.Equal(profile, declaration.Profile);
        Assert.Equal(InvoiceSyntax.Cii, declaration.Syntax);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_EVIDENCE");
        if (output != null) {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, "FacturX-1.09.2-" + profile + ".xml"), xml);
            await File.WriteAllTextAsync(Path.Combine(output, "FacturX-1.09.2-" + profile + ".json"),
                System.Text.Json.JsonSerializer.Serialize(report, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
        }
    }

    [InvoiceStandardsTheory]
    [InlineData(false)]
    public async Task ReleaseProfileMismatchDoesNotClaimThatSchemaValidationRan(bool _) {
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create(), InvoiceTestContracts.En16931());

        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner())
            .ValidateAsync(xml, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31);

        Assert.Equal(InvoiceValidationStatus.NotRun, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.NotRun, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, item => item.Code == "INV-RULESET-PROFILE");
        Assert.DoesNotContain(report.Diagnostics, item => item.Code == "INV-XML");
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, false)]
    [InlineData(InvoiceSyntax.Ubl, false)]
    [InlineData(InvoiceSyntax.Cii, true)]
    [InlineData(InvoiceSyntax.Ubl, true)]
    public async Task RichInvoicesAndCreditNotesPassSchemaRulesAndRoundTrip(InvoiceSyntax syntax, bool credit) {
        Invoice invoice = InvoiceFixture.Rich(credit);
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax));
        var validator = new InvoiceValidator(Bundle(), Runner());
        var report = await validator.ValidateAsync(xml, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(report.IsValid, Report(report));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping, string.Join("\n", read.UnmappedData.Select(d => d.Location + ": " + d.Message)));
        Assert.Equal(xml, read.Write(InvoiceTestContracts.En16931(syntax)));
        var conversion = InvoiceConverter.Convert(xml, InvoiceTestContracts.En16931(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
        Assert.True(conversion.Succeeded, string.Join("\n", conversion.Diagnostics.Select(d => d.Location + ": " + d.Message)));
        var convertedReport = await validator.ValidateAsync(conversion.Xml!, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.True(convertedReport.IsValid, Report(convertedReport));
    }

    [InvoiceStandardsTheory]
    [InlineData("01.01a-INVOICE_ubl.xml")]
    [InlineData("01.01a-INVOICE_uncefact.xml")]
    [InlineData("01.05a-INVOICE_ubl.xml")]
    [InlineData("01.05a-INVOICE_uncefact.xml")]
    public async Task IndependentInvoicesAndTheirConversionsPassTheAuthorityRules(string name) {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "KoSIT", name));
        var validator = new InvoiceValidator(Bundle(), Runner());
        var originalReport = await validator.ValidateAsync(original, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31);
        Assert.True(originalReport.IsValid, Report(originalReport));
        InvoiceReadResult read = InvoiceParser.Read(original);
        var converted = InvoiceConverter.Convert(original, InvoiceTestContracts.XRechnung(read.Declaration.Syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
        Assert.True(converted.Succeeded, string.Join("\n", converted.Diagnostics.Select(d => d.Location + ": " + d.Message)));
        var report = await validator.ValidateAsync(converted.Xml!, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31);
        Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task SchemaOnlyNeverClaimsBusinessRuleSuccessAndWrongAmountsFailRules(InvoiceSyntax syntax) {
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create(), InvoiceTestContracts.En16931(syntax));
        InvoiceRuleBundle bundle = Bundle();
        var partial = await new InvoiceValidator(bundle).ValidateAsync(xml, InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Passed, partial.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.NotRun, partial.BusinessRulesStatus);
        Assert.False(partial.IsValid);
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(xml));
        XElement amount = document.Descendants().Single(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "DuePayableAmount" : "PayableAmount"));
        amount.Value = "0.00";
        var report = await new InvoiceValidator(bundle, Runner()).ValidateAsync(Encoding.UTF8.GetBytes(document.ToString()), InvoiceSpecificationRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, d => d.Code == "BR-CO-16");
    }
    private static string Report(InvoiceValidationReport report) => report.SchemaStatus + "/" + report.BusinessRulesStatus + "\n" + string.Join("\n", report.Diagnostics.Select(d => d.Code + ": " + d.Message));
}
