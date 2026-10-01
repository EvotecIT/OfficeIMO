using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeInvoiceFileWorkflowTests {
    private static InvoiceXmlOptions Target => new(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Ubl, InvoiceProfile.En16931);
    [Fact]
    public async Task SourceHierarchyCollisionPreventsEarlierOutputPublication() {
        using var files = new Files(); string output = Path.Combine(files.Root, "output.xml");
        await Assert.ThrowsAnyAsync<IOException>(() => OfficeInvoiceFileWorkflow.RunBatchAsync([
            new(files.Input, OfficeInvoiceWorkflowOperation.Convert, output, Target),
            new(files.Input, OfficeInvoiceWorkflowOperation.Convert, Path.Combine(files.Input, "nested.xml"), Target)]));
        Assert.False(File.Exists(output));
        Assert.Single(Directory.GetFiles(files.Root));
    }
    [Fact]
    public async Task AncestorOutputCollisionIsRejectedBeforePublication() {
        using var files = new Files();
        string parent = Path.Combine(files.Root, "output"), child = Path.Combine(parent, "invoice.xml");
        foreach (bool reversed in new[] { false, true }) {
            OfficeInvoiceFileWorkflowRequest[] requests = [new(files.Input, OfficeInvoiceWorkflowOperation.Convert, parent, Target), new(files.Input, OfficeInvoiceWorkflowOperation.Convert, child, Target)];
            if (reversed) Array.Reverse(requests);
            await Assert.ThrowsAsync<OfficeInvoiceOutputException>(() => OfficeInvoiceFileWorkflow.RunBatchAsync(requests));
            Assert.False(File.Exists(parent)); Assert.False(Directory.Exists(parent));
        }
    }
    [Fact]
    public async Task OutputBudgetBlocksPublicationWhileRetainingExecutionDiagnostics() {
        using var files = new Files(); string output = Path.Combine(files.Root, "output.xml");
        var result = Assert.Single(await OfficeInvoiceFileWorkflow.RunBatchAsync([new(files.Input, OfficeInvoiceWorkflowOperation.Convert, output, Target)],
            new() { MaximumOutputBytes = 1 }));
        Assert.False(result.Succeeded); Assert.False(result.Published); Assert.Null(result.PublicationError);
        Assert.Contains(result.Workflow.Diagnostics, d => d.Code == "INV-WORKFLOW-BATCH-OUTPUT");
        Assert.False(File.Exists(output)); Assert.Single(Directory.GetFiles(files.Root));
    }
    [Fact]
    public async Task MissingLaterInputPreventsAnyPublication() {
        using var files = new Files(); string output = Path.Combine(files.Root, "output.xml");
        await Assert.ThrowsAsync<FileNotFoundException>(() => OfficeInvoiceFileWorkflow.RunBatchAsync([
            new(files.Input, OfficeInvoiceWorkflowOperation.Convert, output, Target), new(Path.Combine(files.Root, "missing.xml"))]));
        Assert.False(File.Exists(output));
    }
    [Fact]
    public async Task CancelledBatchDoesNotEnumerateInputs() {
        bool enumerated = false;
        IEnumerable<OfficeInvoiceFileWorkflowRequest> Requests() { enumerated = true; yield return new("absent.xml"); }
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeInvoiceFileWorkflow.RunBatchAsync(Requests(), cancellationToken: new(true)));
        Assert.False(enumerated);
    }
    private sealed class Files : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "OfficeIMO-Invoice-Files-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(Root, "source.xml");
        internal Files() {
            Directory.CreateDirectory(Root);
            File.WriteAllBytes(Input, InvoiceSerializer.Write(InvoiceFixture.Create(), new(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931)));
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
