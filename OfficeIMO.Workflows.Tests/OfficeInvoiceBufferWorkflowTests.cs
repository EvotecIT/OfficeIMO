using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public class OfficeInvoiceBufferWorkflowTests {
    private static InvoiceXmlOptions Contract(InvoiceSyntax syntax = InvoiceSyntax.Cii) => new(
        InvoiceSpecificationRelease.En16931_1_3_16, syntax, InvoiceProfile.En16931);
    private static byte[] Xml() => InvoiceSerializer.Write(InvoiceFixture.Create(), Contract());

    [Fact]
    public async Task InspectionCapturesInputAndSeparatesModelChecksFromUnrunStandards() {
        byte[] xml = Xml();
        string hash = Convert.ToHexString(SHA256.HashData(xml));
        var request = new OfficeInvoiceWorkflowRequest(xml, inputName: "invoice.xml");
        Array.Clear(xml);
        byte[] copy = request.ToInputBytes();
        Array.Clear(copy);
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(request);
        Assert.True(result.Succeeded);
        Assert.Equal(hash, result.InputSha256);
        Assert.Equal("invoice.xml", result.InputName);
        Assert.True(result.ModelValidation!.IsValid);
        Assert.Equal(119m, result.ModelValidation.Calculation!.PayableAmount);
        Assert.Equal(InvoiceValidationStatus.NotRun, result.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.NotRun, result.BusinessRulesStatus);
        Assert.Null(result.ToOutputBytes());
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    public async Task AggregateSourceInspectionDoesNotMisreportMissingInvoiceLines(InvoiceProfile profile) {
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii,
            profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        byte[] xml = InvoiceSerializer.Write(InvoiceFixture.Create(), contract);
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(new(xml, OfficeInvoiceWorkflowOperation.Validate));
        Assert.True(result.Succeeded);
        Assert.True(result.ModelValidation!.IsValid);
        Assert.Empty(result.Source!.Invoice.Lines);
        Assert.Equal(119m, result.ModelValidation.Calculation!.PayableAmount);
    }

    [Fact]
    public async Task MalformedInputReportsFailureWithNoOutput() {
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(new(Encoding.UTF8.GetBytes("<bad>")));
        Assert.False(result.Succeeded);
        Assert.Null(result.Source);
        Assert.Contains(result.Diagnostics, d => d.Location == "Source" && d.Code == "INV-WORKFLOW-OPERATION");
        Assert.Null(result.ToOutputBytes());
    }

    [Fact]
    public async Task UnknownSourceDataBlocksConversionButRemainsAvailableForInspection() {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(Xml()));
        document.Root!.Add(new XElement("{urn:example:extension}BusinessFlag", "retain"));
        byte[] xml = Encoding.UTF8.GetBytes(document.ToString());
        OfficeInvoiceWorkflowResult inspected = await OfficeInvoiceBufferWorkflow.RunAsync(new(xml));
        Assert.True(inspected.Succeeded);
        Assert.False(inspected.Source!.HasCompleteMapping);
        OfficeInvoiceWorkflowResult converted = await OfficeInvoiceBufferWorkflow.RunAsync(new(xml, OfficeInvoiceWorkflowOperation.Convert, Contract(InvoiceSyntax.Ubl)));
        Assert.False(converted.Succeeded);
        Assert.Null(converted.ToOutputBytes());
        Assert.Null(converted.ToOutputXmlBytes());
    }

    [Fact]
    public async Task ConversionReturnsIndependentBytesAndProfileProjectionWarnings() {
        var target = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii,
            InvoiceProfile.BasicWithoutLines, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(new(Xml(), OfficeInvoiceWorkflowOperation.Convert, target));
        Assert.True(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-TARGET-PROJECTION" && d.Location == "Lines" && d.Severity == InvoiceDiagnosticSeverity.Warning);
        byte[] output = result.ToOutputBytes()!;
        Assert.Empty(InvoiceParser.Read(output).Invoice.Lines);
        Array.Clear(output);
        Assert.Equal(119m, InvoiceParser.Read(result.ToOutputXmlBytes()!).Invoice.DeclaredTotals!.PayableAmount);
    }

    [Fact]
    public async Task RequestedStandardsValidationCannotSilentlyFallbackToModelChecks() {
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(new(Xml(), OfficeInvoiceWorkflowOperation.Convert,
            Contract(), InvoiceSpecificationRelease.En16931_1_3_16));
        Assert.False(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-WORKFLOW-VALIDATOR");
        Assert.Equal(InvoiceValidationStatus.NotRun, result.SchemaStatus);
        Assert.Null(result.ToOutputBytes());
    }

    [Theory]
    [InlineData(OfficeInvoiceWorkflowOperation.RenderPresentationPdf)]
    [InlineData(OfficeInvoiceWorkflowOperation.RenderHybridPdf)]
    public async Task RenderedPdfAndXmlUseOneCapturedModel(OfficeInvoiceWorkflowOperation operation) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "SourceSerif4-Regular.otf"));
        var options = new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Source Serif");
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931);
        OfficeInvoiceWorkflowResult result = await OfficeInvoiceBufferWorkflow.RunAsync(new(Xml(), operation, contract, pdfOptions: options));
        Assert.True(result.Succeeded, string.Join("; ", result.Diagnostics.Select(d => d.Message)));
        byte[] pdf = result.ToOutputBytes()!;
        Assert.Contains("119.00 EUR", PdfReadDocument.Open(pdf).ExtractText());
        Assert.Equal(119m, InvoiceParser.Read(result.ToOutputXmlBytes()!).Invoice.DeclaredTotals!.PayableAmount);
        if (operation == OfficeInvoiceWorkflowOperation.RenderHybridPdf)
            Assert.Equal(result.ToOutputXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        else Assert.Empty(PdfDocument.Load(pdf).Attachments.Extract());
    }

    [Fact]
    public async Task BatchPreflightBoundsLazyEnumerationAndCombinedInput() {
        var request = new OfficeInvoiceWorkflowRequest(Xml());
        int yielded = 0;
        IEnumerable<OfficeInvoiceWorkflowRequest> Unbounded() { while (true) { yielded++; yield return request; } }
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeInvoiceBufferWorkflow.RunBatchAsync(Unbounded(), new() { MaximumRequests = 2 }));
        Assert.Equal(3, yielded);
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeInvoiceBufferWorkflow.RunBatchAsync(new[] { request }, new() { MaximumInputBytes = 1 }));
    }

    [Fact]
    public async Task BatchOutputLimitDropsTheArtifactAndFailurePolicyControlsContinuation() {
        var valid = new OfficeInvoiceWorkflowRequest(Xml(), OfficeInvoiceWorkflowOperation.Convert, Contract(InvoiceSyntax.Ubl));
        IReadOnlyList<OfficeInvoiceWorkflowResult> bounded = await OfficeInvoiceBufferWorkflow.RunBatchAsync(new[] { valid }, new() { MaximumOutputBytes = 1 });
        OfficeInvoiceWorkflowResult failed = Assert.Single(bounded);
        Assert.False(failed.Succeeded);
        Assert.Contains(failed.Diagnostics, d => d.Code == "INV-WORKFLOW-BATCH-OUTPUT");
        Assert.Null(failed.ToOutputBytes());
        Assert.Null(failed.ToOutputXmlBytes());
        var malformed = new OfficeInvoiceWorkflowRequest(Encoding.UTF8.GetBytes("<bad>"));
        Assert.Single(await OfficeInvoiceBufferWorkflow.RunBatchAsync(new[] { malformed, valid }, new() { ContinueOnFailure = false }));
        IReadOnlyList<OfficeInvoiceWorkflowResult> continued = await OfficeInvoiceBufferWorkflow.RunBatchAsync(new[] { malformed, valid });
        Assert.Equal(2, continued.Count);
        Assert.False(continued[0].Succeeded);
        Assert.True(continued[1].Succeeded);
    }

    [Fact]
    public async Task CancellationStopsBeforeEnumeratingBatchInput() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        int enumerated = 0;
        IEnumerable<OfficeInvoiceWorkflowRequest> Inputs() { enumerated++; yield return new(Xml()); }
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeInvoiceBufferWorkflow.RunBatchAsync(Inputs(), cancellationToken: cancellation.Token));
        Assert.Equal(0, enumerated);
    }
}
