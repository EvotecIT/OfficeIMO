using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeInvoiceStorageWorkflowTests {
    private static byte[] Xml() => InvoiceSerializer.Write(InvoiceFixture.Create(),
        new(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.En16931));

    [Fact]
    public async Task LocalEditingUsesVerifiedSourcesAndKeepsExistingOutputsWithRename() {
        using var files = new Files();
        string input = Path.Combine(files.Root, "invoice.xml"), output = Path.Combine(files.Root, "edited.xml");
        byte[] original = Xml(); File.WriteAllBytes(input, original); File.WriteAllText(output, "Existing output");
        var result = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
            InputPath = input, Operation = OfficeInvoiceWorkflowOperation.EditSource, SourceEdits = new(number: "EDITED"),
            OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
        });
        Assert.True(result.Succeeded, result.Summary);
        Assert.NotEqual(output, result.OutputPath);
        Assert.Equal("Existing output", File.ReadAllText(output));
        Assert.Equal(original, File.ReadAllBytes(input));
        Assert.Equal("EDITED", InvoiceSourceDocument.Load(result.OutputPath!).Number);
        Assert.Contains(result.Diagnostics, d => d.Code == "AtomicPublication");
        Assert.Equal(3, Directory.GetFiles(files.Root).Length);
    }

    [Theory]
    [InlineData("success")]
    [InlineData("write-failure")]
    [InlineData("source-change")]
    [InlineData("denied")]
    [InlineData("deferred-alias")]
    public async Task ProviderXmlPublicationUsesRecoverySourceChecksAndActualDestinationAuthorization(string mode) {
        using var files = new Files();
        var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(files.Root, "recovery"));
        byte[] source = Xml(), stored = Xml(); int writes = 0, reads = 0;
        var output = new OfficeWorkflowStreamOutput("edited.xml", _ => Task.FromResult<Stream>(new MemoryStream(stored)), _ => {
            writes++;
            Assert.NotEmpty(Directory.GetFiles(recovery.DirectoryPath, "record.json", SearchOption.AllDirectories));
            if (mode == "write-failure") throw new IOException("Provider write failed.");
            return Task.FromResult<Stream>(new CommitStream(bytes => stored = bytes));
        }, recovery, mode == "deferred-alias" ? _ => Task.FromResult("content://invoice/source") : null);
        var result = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
            InputPath = "content://invoice/source", InputStream = new("invoice.xml", _ => { reads++; return Task.FromResult<Stream>(new MemoryStream(source)); }),
            Operation = OfficeInvoiceWorkflowOperation.EditSource, SourceEdits = new(number: "EDITED"),
            OutputPath = "content://invoice/destination", OutputStream = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            PublicationGuard = new Guard(() => {
                if (mode == "source-change") source = System.Text.Encoding.UTF8.GetBytes("<changed/>");
                return mode != "denied";
            })
        });
        Assert.True(result.Workflow!.Succeeded);
        Assert.Equal(mode == "success" ? OfficeWorkflowStatus.Completed : mode is "write-failure" or "deferred-alias"
            ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(mode is "success" or "write-failure" ? 1 : 0, writes);
        if (mode == "success") {
            Assert.True(result.Succeeded); Assert.Equal("EDITED", InvoiceSourceDocument.Load(stored).Number);
            Assert.Equal("content://invoice/destination", result.OutputPath); Assert.True(reads >= 2);
            Assert.Empty(recovery.GetRecoveries());
        } else {
            Assert.False(result.Succeeded); Assert.Null(result.OutputPath);
            Assert.Equal("INV-2026-001", InvoiceSourceDocument.Load(stored).Number);
            if (result.Status == OfficeWorkflowStatus.Unconfirmed) {
                var retained = Assert.Single(recovery.GetRecoveries());
                Assert.Equal(retained.Id, result.Recovery!.Id); Assert.Equal(retained.FilePath, result.Recovery.FilePath);
                await recovery.VerifyAsync(retained);
                Assert.Equal("EDITED", InvoiceSourceDocument.Load(retained.FilePath).Number);
            } else { Assert.Empty(recovery.GetRecoveries()); Assert.Null(result.Recovery); }
        }
    }

    [Fact]
    public async Task CancellationAndInvalidDestinationContractDoNotAcquireProviderInput() {
        int opens = 0;
        var input = new OfficeWorkflowStreamInput("invoice.xml", _ => { opens++; return Task.FromResult<Stream>(new MemoryStream(Xml())); });
        var cancelled = await new OfficeWorkflowRunner().RunInvoiceAsync(new() { InputPath = "content://invoice/input", InputStream = input }, cancellationToken: new(true));
        Assert.Equal(OfficeWorkflowStatus.Cancelled, cancelled.Status); Assert.Null(cancelled.Workflow);
        var invalid = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
            InputPath = "content://invoice/input", InputStream = input, Operation = OfficeInvoiceWorkflowOperation.EditSource,
            SourceEdits = new(number: "EDITED"), OutputPath = Path.Combine(Path.GetTempPath(), "wrong.pdf")
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, invalid.Status); Assert.Null(invalid.Workflow); Assert.Equal(0, opens);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task InputAndOutputLimitsPreventPublication(bool inputLimit) {
        using var files = new Files();
        string input = Path.Combine(files.Root, "invoice.xml"), output = Path.Combine(files.Root, "edited.xml");
        File.WriteAllBytes(input, Xml());
        var result = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
            InputPath = input, Operation = OfficeInvoiceWorkflowOperation.EditSource, SourceEdits = new(number: "EDITED"), OutputPath = output,
            MaximumInputBytes = inputLimit ? 1 : InvoiceProfileDeclaration.MaximumXmlBytes, MaximumOutputBytes = inputLimit ? 65536 : 1
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status); Assert.False(File.Exists(output));
        Assert.Single(Directory.GetFiles(files.Root));
        if (!inputLimit) Assert.Contains(result.Workflow!.Diagnostics, d => d.Code == "INV-WORKFLOW-BATCH-OUTPUT");
    }

    [Fact]
    public async Task OutputCannotReplaceItsInputAndRequestedValidationCannotFallback() {
        using var files = new Files(); string input = Path.Combine(files.Root, "invoice.xml"), output = Path.Combine(files.Root, "edited.xml");
        byte[] original = Xml(); File.WriteAllBytes(input, original);
        var collision = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
            InputPath = input, Operation = OfficeInvoiceWorkflowOperation.EditSource, SourceEdits = new(number: "EDITED"),
            OutputPath = input, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });
        Assert.False(collision.Succeeded); Assert.Equal(original, File.ReadAllBytes(input));
        var unvalidated = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
            InputPath = input, Operation = OfficeInvoiceWorkflowOperation.EditSource, SourceEdits = new(number: "EDITED"), OutputPath = output,
            ValidationRelease = InvoiceSpecificationRelease.En16931_1_3_16
        });
        Assert.False(unvalidated.Succeeded); Assert.False(File.Exists(output));
        Assert.Contains(unvalidated.Workflow!.Diagnostics, d => d.Code == "INV-WORKFLOW-VALIDATOR");
    }

    private sealed class Guard(Func<bool> action) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) => ValueTask.FromResult(action());
    }
    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) { if (disposing && !_closed) { _closed = true; commit(ToArray()); } base.Dispose(disposing); }
    }
    private sealed class Files : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "OfficeIMO-Invoice-Storage-" + Guid.NewGuid().ToString("N"));
        internal Files() => Directory.CreateDirectory(Root);
        public void Dispose() => Directory.Delete(Root, true);
    }
}
