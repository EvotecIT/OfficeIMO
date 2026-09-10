using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public sealed class InvoiceStandardsTheoryAttribute : TheoryAttribute {
    public InvoiceStandardsTheoryAttribute() {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_STANDARDS_TESTS") != "1") Skip = "Opt-in authority artifact validation; run Build/Test-InvoicingStandards.ps1.";
    }
}
public class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, true)]
    [InlineData(InvoiceSyntax.Ubl, true)]
    [InlineData(InvoiceSyntax.Cii, false)]
    [InlineData(InvoiceSyntax.Ubl, false)]
    public async Task MultipleAccountsWithSingletonPaymentDetailsPassAuthorityRules(InvoiceSyntax syntax, bool card) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.Accounts.Add(new InvoiceBankAccount { Identifier = "DE89370400440532013000", Name = "Second account" });
        invoice.Payment.MeansCode = card ? "48" : "59";
        if (card) { invoice.Payment.CardNumber = "1234"; invoice.Payment.CardHolder = "Card Holder"; }
        else { invoice.Payment.MandateReference = "mandate-1"; invoice.Payment.DebitedAccount = "DE89370400440532013000"; invoice.Payment.CreditorIdentifier = "DE98ZZZ09999999999"; }
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, InvoiceRulesRelease.En16931_1_3_16);
        Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ExcessiveContentFailuresRemainInvalidInsteadOfEngineFailure(bool schemaErrors) {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Create(), new InvoiceXmlOptions(InvoiceSyntax.Ubl))));
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
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, InvoiceRulesRelease.En16931_1_3_16);
        Assert.Equal(schemaErrors ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(schemaErrors ? InvoiceValidationStatus.NotRun : InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, d => d.Code == "INV-DIAGNOSTICS-TRUNCATED" && d.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.DoesNotContain(report.Diagnostics, d => d.Code == "INV-SCHEMA-ENGINE" || d.Code == "INV-RULES-ENGINE");
        Assert.InRange(report.Diagnostics.Count, 1, 1000);
    }

    private static InvoiceRuleBundle Bundle() => InvoiceRuleBundle.Load(
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE") ?? throw new InvalidOperationException("Rule bundle path is required."),
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PEPPOL_RULES"));
    private static SaxonInvoiceRulesRunner Runner() => new SaxonInvoiceRulesRunner(
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR") ?? throw new InvalidOperationException("Saxon path is required."),
        Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java");

    [InvoiceStandardsTheory]
    [InlineData(InvoiceRulesRelease.En16931_1_3_16)]
    public async Task MissingRuntimeIsAnEngineFailureAndCancellationIsPreserved(InvoiceRulesRelease release) {
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create());
        string jar = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!;
        var validator = new InvoiceValidator(Bundle(), new SaxonInvoiceRulesRunner(jar, "officeimo-nonexistent-java-" + Guid.NewGuid().ToString("N")));
        InvoiceValidationReport report = await validator.ValidateAsync(xml, release);
        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Failed, report.BusinessRulesStatus);
        Assert.False(report.IsValid);
        Assert.Contains(report.Diagnostics, d => d.Code == "INV-RULES-ENGINE");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => validator.ValidateAsync(xml, release, cancelled.Token));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.En16931, InvoiceRulesRelease.En16931_1_3_16)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.En16931, InvoiceRulesRelease.En16931_1_3_16)]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung, InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung, InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, InvoiceRulesRelease.PeppolBis_3_0_21)]
    public async Task GeneratedInvoicesPassPinnedSchemaAndBusinessRules(InvoiceSyntax syntax, InvoiceProfile profile, InvoiceRulesRelease release) {
        var invoice = InvoiceFixture.Create();
        if (profile == InvoiceProfile.PeppolBis) {
            invoice.Seller.ElectronicAddress = new InvoiceIdentifier("1234567890128", "0088");
            invoice.Buyer.ElectronicAddress = new InvoiceIdentifier("1234567890135", "0088");
        }
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax, profile));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, release);
        Assert.True(report.IsValid, Report(report));
        Assert.Equal(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(xml)), report.Sha256);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_EVIDENCE");
        if (output != null) {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, syntax + "-" + profile + ".xml"), xml);
            await File.WriteAllTextAsync(Path.Combine(output, syntax + "-" + profile + ".json"), System.Text.Json.JsonSerializer.Serialize(report, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
        }
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, false)]
    [InlineData(InvoiceSyntax.Ubl, false)]
    [InlineData(InvoiceSyntax.Cii, true)]
    [InlineData(InvoiceSyntax.Ubl, true)]
    public async Task RichInvoicesAndCreditNotesPassSchemaRulesAndRoundTrip(InvoiceSyntax syntax, bool credit) {
        Invoice invoice = InvoiceFixture.Rich(credit);
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        var validator = new InvoiceValidator(Bundle(), Runner());
        var report = await validator.ValidateAsync(xml, InvoiceRulesRelease.En16931_1_3_16);
        Assert.True(report.IsValid, Report(report));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping, string.Join("\n", read.UnmappedData.Select(d => d.Location + ": " + d.Message)));
        Assert.Equal(xml, read.Write());
        var conversion = InvoiceConverter.Convert(xml, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
        Assert.True(conversion.Succeeded, string.Join("\n", conversion.Diagnostics.Select(d => d.Location + ": " + d.Message)));
        var convertedReport = await validator.ValidateAsync(conversion.Xml!, InvoiceRulesRelease.En16931_1_3_16);
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
        var originalReport = await validator.ValidateAsync(original, InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31);
        Assert.True(originalReport.IsValid, Report(originalReport));
        InvoiceReadResult read = InvoiceParser.Read(original);
        var converted = InvoiceConverter.Convert(original, new InvoiceXmlOptions(read.Declaration.Syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii, InvoiceProfile.XRechnung));
        Assert.True(converted.Succeeded, string.Join("\n", converted.Diagnostics.Select(d => d.Location + ": " + d.Message)));
        var report = await validator.ValidateAsync(converted.Xml!, InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31);
        Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task SchemaOnlyNeverClaimsBusinessRuleSuccessAndWrongAmountsFailRules(InvoiceSyntax syntax) {
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create(), new InvoiceXmlOptions(syntax));
        InvoiceRuleBundle bundle = Bundle();
        var partial = await new InvoiceValidator(bundle).ValidateAsync(xml, InvoiceRulesRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Passed, partial.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.NotRun, partial.BusinessRulesStatus);
        Assert.False(partial.IsValid);
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(xml));
        XElement amount = document.Descendants().Single(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "DuePayableAmount" : "PayableAmount"));
        amount.Value = "0.00";
        var report = await new InvoiceValidator(bundle, Runner()).ValidateAsync(Encoding.UTF8.GetBytes(document.ToString()), InvoiceRulesRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        Assert.Contains(report.Diagnostics, d => d.Code == "BR-CO-16");
    }
    private static string Report(InvoiceValidationReport report) => report.SchemaStatus + "/" + report.BusinessRulesStatus + "\n" + string.Join("\n", report.Diagnostics.Select(d => d.Code + ": " + d.Message));
}
