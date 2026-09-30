using System.Security.Cryptography;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Invoicing.Validation;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class InvoiceWorkflowStandardsTheoryAttribute : TheoryAttribute {
    public InvoiceWorkflowStandardsTheoryAttribute() {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_STANDARDS_TESTS") != "1")
            Skip = "Requires explicitly configured pinned invoice artifacts and Saxon runtime.";
    }
}

public class OfficeInvoiceStandardsWorkflowTests {
    [InvoiceWorkflowStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, false)]
    [InlineData(InvoiceSyntax.Ubl, false)]
    [InlineData(InvoiceSyntax.Cii, true)]
    [InlineData(InvoiceSyntax.Ubl, true)]
    public async Task SourceEditsValidateExactPreservedOutputAndBlockInvalidExtensions(InvoiceSyntax syntax, bool extension) {
        var bundle = InvoiceRuleBundle.Load(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE")!);
        var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!,
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java");
        var validator = new InvoiceValidator(bundle, runner);
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.En16931_1_3_16, syntax, InvoiceProfile.En16931);
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create(), contract);
        if (extension) {
            var document = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(xml));
            document.Root!.Add(new System.Xml.Linq.XElement("{urn:example:extension}Unmapped", "retain"));
            xml = System.Text.Encoding.UTF8.GetBytes(document.ToString());
        }
        var result = await OfficeInvoiceBufferWorkflow.RunAsync(OfficeInvoiceWorkflowRequest.ForSourceEdit(xml,
            new(number: "EDITED-2026-123", buyerReference: "EDITED-BUYER"), contract.Release), validator);
        Assert.Equal(!extension, result.Succeeded);
        Assert.NotEqual(result.InputSha256, result.StandardsValidation!.Sha256);
        if (extension) {
            Assert.Equal(InvoiceValidationStatus.Invalid, result.SchemaStatus);
            Assert.Null(result.ToOutputBytes());
            Assert.Contains(result.Diagnostics, d => d.Code == "INV-WORKFLOW-STANDARDS");
        } else {
            Assert.Equal(InvoiceValidationStatus.Passed, result.SchemaStatus);
            Assert.Equal(InvoiceValidationStatus.Passed, result.BusinessRulesStatus);
            Assert.Equal(Convert.ToHexString(SHA256.HashData(result.ToOutputBytes()!)), result.StandardsValidation.Sha256);
            Assert.Equal("EDITED-2026-123", InvoiceSourceDocument.Load(result.ToOutputBytes()!).Number);
        }
    }

    [InvoiceWorkflowStandardsTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task CompletedInvalidStandardsInspectionStillReportsAndContinues(bool schemaInvalid) {
        var bundle = InvoiceRuleBundle.Load(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE")!);
        var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!,
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java");
        var validator = new InvoiceValidator(bundle, runner);
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931);
        var invoice = InvoiceFixture.Create();
        if (!schemaInvalid) invoice.TypeCode = "999";
        byte[] xml = InvoiceSerializer.Write(invoice, contract);
        if (schemaInvalid) {
            var document = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(xml));
            document.Root!.Add(new System.Xml.Linq.XElement("{urn:example:extension}Unmapped", "retain"));
            xml = System.Text.Encoding.UTF8.GetBytes(document.ToString());
        }
        var inspect = new OfficeInvoiceWorkflowRequest(xml, validationRelease: contract.Release);
        var results = await OfficeInvoiceBufferWorkflow.RunBatchAsync([inspect, inspect], new() { ContinueOnFailure = false }, validator);
        Assert.Equal(2, results.Count);
        Assert.All(results, result => {
            Assert.True(result.Succeeded);
            Assert.False(result.StandardsValidation!.IsValid);
            Assert.Equal(schemaInvalid ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed, result.SchemaStatus);
            Assert.Equal(schemaInvalid ? InvoiceValidationStatus.NotRun : InvoiceValidationStatus.Invalid, result.BusinessRulesStatus);
            Assert.Contains(result.Diagnostics, d => d.Severity == InvoiceDiagnosticSeverity.Error);
        });
        var validated = await OfficeInvoiceBufferWorkflow.RunAsync(new(xml, OfficeInvoiceWorkflowOperation.Validate, validationRelease: contract.Release), validator);
        Assert.False(validated.Succeeded);
        Assert.Null(validated.ToOutputBytes());
    }
    [InvoiceWorkflowStandardsTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task StandardsEvidenceMatchesConvertedBytesAndIncompleteRulesBlockOutput(bool runRules) {
        var bundle = InvoiceRuleBundle.Load(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_RULE_BUNDLE")!,
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PEPPOL_RULES"));
        var runner = runRules ? new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!,
            Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java") : null;
        var validator = new InvoiceValidator(bundle, runner);
        var sourceContract = new InvoiceXmlOptions(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Ubl, InvoiceProfile.En16931);
        var target = new InvoiceXmlOptions(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931);
        byte[] source = InvoiceSerializer.Write(InvoiceFixture.Create(), sourceContract);
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(new(source, OfficeInvoiceWorkflowOperation.Convert,
            target, target.Release), validator);
        Assert.Equal(InvoiceValidationStatus.Passed, result.SchemaStatus);
        Assert.Equal(runRules, result.Succeeded);
        if (runRules) {
            Assert.Equal(InvoiceValidationStatus.Passed, result.BusinessRulesStatus);
            Assert.Equal(Convert.ToHexString(SHA256.HashData(result.ToOutputBytes()!)), result.StandardsValidation!.Sha256);
            Assert.NotEqual(result.InputSha256, result.StandardsValidation.Sha256);
            Assert.Equal(result.ToOutputBytes(), result.ToOutputXmlBytes());
        } else {
            Assert.Equal(InvoiceValidationStatus.NotRun, result.BusinessRulesStatus);
            Assert.Null(result.ToOutputBytes());
            Assert.Contains(result.Diagnostics, d => d.Code == "INV-WORKFLOW-STANDARDS");
        }
    }
}
