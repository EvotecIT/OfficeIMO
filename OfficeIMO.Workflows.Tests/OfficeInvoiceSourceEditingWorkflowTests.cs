using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Invoicing.Validation;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeInvoiceSourceEditingWorkflowTests {
    private static readonly XNamespace Extension = "urn:example:preserved";
    private static byte[] Source(InvoiceSyntax syntax) {
        byte[] bytes = InvoiceSerializer.Write(InvoiceFixture.Create(), new(InvoiceSpecificationRelease.En16931_1_3_16, syntax, InvoiceProfile.En16931));
        var document = XDocument.Parse(Encoding.UTF8.GetString(bytes));
        document.Root!.Add(new XElement(Extension + "Payload", new XAttribute("flag", "retain"), "business data"));
        return Encoding.UTF8.GetBytes(document.ToString());
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task SourceEditRetainsExtensionsAndReportsMappingSeparatelyFromCompletion(InvoiceSyntax syntax) {
        byte[] bytes = Source(syntax);
        string hash = Convert.ToHexString(SHA256.HashData(bytes));
        var request = OfficeInvoiceWorkflowRequest.ForSourceEdit(bytes, new(number: "EDITED-123", buyerReference: "EDITED-BUYER"));
        Array.Clear(bytes);
        var result = await OfficeInvoiceBufferWorkflow.RunAsync(request);
        Assert.True(result.Succeeded);
        Assert.Equal(hash, result.InputSha256);
        Assert.False(result.Source!.HasCompleteMapping);
        Assert.Equal("EDITED-123", result.Source.Invoice.Number);
        Assert.Equal("EDITED-BUYER", result.Source.Invoice.BuyerReference);
        Assert.Equal(119m, result.ModelValidation!.Calculation!.PayableAmount);
        byte[] output = result.ToOutputBytes()!;
        Assert.Equal(output, result.ToOutputXmlBytes());
        Assert.Equal("business data", XDocument.Parse(Encoding.UTF8.GetString(output)).Root!.Element(Extension + "Payload")!.Value);
        Array.Clear(output);
        Assert.Equal("EDITED-123", InvoiceSourceDocument.Load(result.ToOutputBytes()!).Number);
        Assert.Equal(InvoiceValidationStatus.NotRun, result.SchemaStatus);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-SOURCE-EDIT-VALIDATION-REQUIRED");
    }

    [Fact]
    public async Task ExistingInvalidBusinessDataCanBeRetainedWithoutClaimingModelValidity() {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(Source(InvoiceSyntax.Ubl)));
        XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
        document.Root!.Element(cac + "AccountingSupplierParty")!.Remove();
        var result = await OfficeInvoiceBufferWorkflow.RunAsync(OfficeInvoiceWorkflowRequest.ForSourceEdit(Encoding.UTF8.GetBytes(document.ToString()), new(number: "EDITED")));
        Assert.True(result.Succeeded);
        Assert.False(result.ModelValidation!.IsValid);
        Assert.Contains(result.Diagnostics, d => d.Location == "Seller" && d.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.NotNull(result.ToOutputBytes());
    }

    [Fact]
    public async Task SignedSourceAndUnavailableRequestedValidatorReturnNoArtifact() {
        byte[] xml = Source(InvoiceSyntax.Ubl);
        var document = XDocument.Parse(Encoding.UTF8.GetString(xml));
        document.Root!.Add(new XElement("{http://www.w3.org/2000/09/xmldsig#}Signature"));
        var signed = await OfficeInvoiceBufferWorkflow.RunAsync(OfficeInvoiceWorkflowRequest.ForSourceEdit(Encoding.UTF8.GetBytes(document.ToString()), new(number: "EDITED")));
        Assert.False(signed.Succeeded); Assert.Null(signed.ToOutputBytes());
        Assert.Contains(signed.Diagnostics, d => d.Location == "Source.Signature");
        var missing = await OfficeInvoiceBufferWorkflow.RunAsync(OfficeInvoiceWorkflowRequest.ForSourceEdit(xml, new(number: "EDITED"), InvoiceSpecificationRelease.En16931_1_3_16));
        Assert.False(missing.Succeeded); Assert.Null(missing.ToOutputXmlBytes());
        Assert.Contains(missing.Diagnostics, d => d.Code == "INV-WORKFLOW-VALIDATOR");
    }

    [Fact]
    public async Task FileEditingUsesSourceProtectionAndAtomicNewFilePublication() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Source-Edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "input.xml"), output = Path.Combine(root, "edited.xml");
            byte[] bytes = Source(InvoiceSyntax.Cii); File.WriteAllBytes(input, bytes);
            var result = await OfficeInvoiceFileWorkflow.RunAsync(OfficeInvoiceFileWorkflowRequest.ForSourceEdit(input, output, new(number: "EDITED")));
            Assert.True(result.Succeeded); Assert.True(result.Published);
            Assert.Equal(bytes, File.ReadAllBytes(input));
            Assert.Equal("EDITED", InvoiceSourceDocument.Load(output).Number);
            await Assert.ThrowsAsync<OfficeInvoiceOutputException>(() => OfficeInvoiceFileWorkflow.RunAsync(OfficeInvoiceFileWorkflowRequest.ForSourceEdit(input, input, new(number: "NO"))));
            await Assert.ThrowsAsync<OfficeInvoiceOutputException>(() => OfficeInvoiceFileWorkflow.RunAsync(OfficeInvoiceFileWorkflowRequest.ForSourceEdit(input, output, new(number: "NO"))));
            Assert.Equal("EDITED", InvoiceSourceDocument.Load(output).Number);
            Assert.Equal(2, Directory.GetFiles(root).Length);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task EditedXmlCountsOnceAgainstBatchArtifactBudget() {
        var request = OfficeInvoiceWorkflowRequest.ForSourceEdit(Source(InvoiceSyntax.Ubl), new(number: "EDITED"));
        var first = await OfficeInvoiceBufferWorkflow.RunAsync(request);
        var exact = Assert.Single(await OfficeInvoiceBufferWorkflow.RunBatchAsync([request], new() { MaximumOutputBytes = first.OutputByteLength }));
        Assert.True(exact.Succeeded);
        var tooSmall = Assert.Single(await OfficeInvoiceBufferWorkflow.RunBatchAsync([request], new() { MaximumOutputBytes = first.OutputByteLength - 1 }));
        Assert.False(tooSmall.Succeeded); Assert.Null(tooSmall.ToOutputBytes());
        Assert.Contains(tooSmall.Diagnostics, d => d.Code == "INV-WORKFLOW-BATCH-OUTPUT");
    }

    [Fact]
    public void SourceEditCannotBeConstructedAsAnUncapturedOrRetargetedOperation() {
        Assert.Throws<ArgumentException>(() => new OfficeInvoiceWorkflowRequest(Source(InvoiceSyntax.Ubl), OfficeInvoiceWorkflowOperation.EditSource));
        Assert.Throws<ArgumentNullException>(() => OfficeInvoiceWorkflowRequest.ForSourceEdit(Source(InvoiceSyntax.Ubl), null!));
    }
}
