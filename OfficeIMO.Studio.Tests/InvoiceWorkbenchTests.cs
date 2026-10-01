using Avalonia;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class InvoiceWorkbenchTests {
    [Fact]
    public async Task SourceEditingAndPresentationUseRealSharedOwnersAndKeepSourceAndReferences() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string input = Path.Combine(services.Paths.Root, "source.xml");
            byte[] original = InvoiceSample.Create(); File.WriteAllBytes(input, original);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            shell.Commands["Invoices"].Execute(null);
            Assert.True(shell.IsInvoiceMode); Assert.False(shell.HasDocument);
            var model = shell.InvoiceWorkbench; model.InputPath = input;
            model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.EditSource);
            model.EditNumber = "EDITED-42"; model.EditDueDate = "2026-11-30";
            await model.RunCommand.ExecuteAsync(null);
            Assert.True(model.HasOutput, model.Status);
            var edited = InvoiceSourceDocument.Load(model.OutputPath!);
            Assert.Equal("EDITED-42", edited.Number); Assert.Equal(new DateTime(2026, 11, 30), edited.DueDate);
            Assert.Equal("INVOICE-2026-0042", edited.PaymentReference);
            Assert.Equal(original, File.ReadAllBytes(input));
            Assert.Equal("Not requested", model.SchemaSummary); Assert.Equal("Not requested", model.RulesSummary);
            Assert.Contains(model.Diagnostics, finding => finding.Contains("INV-SOURCE-EDIT-VALIDATION-REQUIRED", StringComparison.Ordinal));
            model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.RenderPresentationPdf);
            await model.RunCommand.ExecuteAsync(null);
            Assert.True(model.HasOutput, model.Status);
            Assert.NotEmpty(PdfDocument.Load(model.OutputPath!).Inspect().Pages);
            Assert.Equal(2, services.Jobs.Entries.Count);
            Assert.All(services.Jobs.Entries, job => Assert.Equal(OfficeWorkflowStatus.Completed, job.Outcome));
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderInputIsClosedAndRequiresExplicitOutputForWriting(bool denied) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var file = new TestStorageFile("content://invoice/input", InvoiceSample.Create(), "Selected invoice.xml") { DenyRead = denied };
            string location = await services.Storage.RegisterAsync(file.Item, default);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var model = shell.InvoiceWorkbench; model.InputPath = location;
            await model.RunCommand.ExecuteAsync(null);
            Assert.False(model.HasOutput); Assert.Equal(denied ? 0 : file.Reads, file.ClosedReads);
            Assert.Equal(0, file.Writes);
            if (denied) Assert.Contains(model.Diagnostics, finding => finding.Contains("permission", StringComparison.OrdinalIgnoreCase));
            else Assert.Contains("INVOICE-2026-0042", model.SourceSummary);
            model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.EditSource);
            model.EditNumber = "EDITED"; Assert.False(model.CanRun);
            model.OutputFolder = services.Paths.Root; Assert.True(model.CanRun);
            await model.RunCommand.ExecuteAsync(null);
            Assert.Equal(!denied, model.HasOutput);
            if (!denied) Assert.Equal("EDITED", InvoiceSourceDocument.Load(model.OutputPath!).Number);
            Assert.Equal(0, file.Writes);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData("date")]
    [InlineData("empty")]
    [InlineData("standards")]
    public async Task InvalidOptionsFailBeforeRunnerAndDoNotExposeEarlierOutput(string invalid) {
        var runner = new RecordingRunner();
        using var model = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null), runner) {
            InputPath = Path.Combine(Path.GetTempPath(), "input.xml"), OutputPath = "previous.xml"
        };
        model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.EditSource);
        if (invalid == "date") model.EditDueDate = "30/11/2026";
        if (invalid == "standards") { model.EditNumber = "EDITED"; model.RequireStandards = true; }
        await model.RunCommand.ExecuteAsync(null);
        Assert.Null(runner.Request); Assert.False(model.HasOutput); Assert.False(model.IsBusy);
        Assert.NotEmpty(model.Status);
    }

    [Fact]
    public async Task CancellationReachesInvoiceRunnerAndCapturesRenderingChoicesBeforeAwait() {
        var runner = new RecordingRunner();
        using var model = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null), runner) {
            InputPath = Path.Combine(Path.GetTempPath(), "input.xml")
        };
        model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.RenderPresentationPdf);
        model.SelectedTarget = model.Targets.First(t => t.Options.Syntax == InvoiceSyntax.Cii);
        var running = model.RunCommand.ExecuteAsync(null);
        await runner.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(model.CanCancel); Assert.False(model.CanEdit);
        var columns = runner.Request!.Layout.LineColumns.ToArray();
        model.SelectedTableLayout = model.TableLayouts[0]; model.ModernLayout = false; model.InputPath = "changed.xml";
        Assert.Equal(columns, runner.Request.Layout.LineColumns);
        Assert.EndsWith("input.xml", runner.Request.InputPath);
        model.CancelCommand.Execute(null);
        await running.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(runner.Token.IsCancellationRequested); Assert.False(model.IsBusy); Assert.False(model.HasOutput);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderPublicationUsesConfirmedCapturedEditsAndRecoveryBeforeCreatingAFile(bool consent) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.xml");
            byte[] original = InvoiceSample.Create(); File.WriteAllBytes(source, original);
            var folder = new StudioProviderOutputFolderTests.OutputFolder {
                BeforeCreate = () => Assert.NotEmpty(Directory.GetFiles(services.WorkflowRecovery.DirectoryPath, "record.json", SearchOption.AllDirectories))
            };
            string location = (await services.Storage.RegisterFolderAsync([folder.Item], default))!;
            InvoiceWorkbenchViewModel? model = null;
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                confirmProviderWrite: _ => {
                    // A confirmation callback can yield while other host code changes observable properties.
                    model!.EditNumber = "CHANGED-DURING-CONSENT"; model.InputPath = "changed.xml";
                    model.OutputFolder = "changed-folder"; model.SelectedOperation = model.Operations[0];
                    return Task.FromResult(consent);
                });
            model = shell.InvoiceWorkbench; model.InputPath = source; model.OutputFolder = location;
            model.SelectedOperation = model.Operations.Single(o => o.Value == OfficeInvoiceWorkflowOperation.EditSource);
            model.EditNumber = "CAPTURED";
            await model.RunCommand.ExecuteAsync(null);
            Assert.Equal(consent ? 1 : 0, folder.Creations); Assert.Equal(consent, model.HasOutput);
            if (consent) {
                var file = Assert.Single(folder.Files).Value;
                Assert.Equal("CAPTURED", InvoiceSourceDocument.Load(file.Bytes).Number);
                Assert.Equal("source.edited.xml", file.Name); Assert.Equal(file.Location.AbsoluteUri, model.OutputPath);
                Assert.Equal(OfficeWorkflowStatus.Completed, Assert.Single(services.Jobs.Entries).Outcome);
            } else Assert.Empty(services.Jobs.Entries);
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.Empty(services.WorkflowRecovery.GetRecoveries());
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false, "Not requested")]
    [InlineData(true, "Not run")]
    public void UnexecutedStandardsStageRetainsCapturedRequestIntent(bool requested, string expected) {
        using var model = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null));
        model.RequireStandards = !requested;
        Assert.Equal(expected, model.FormatStandardsStatus(InvoiceValidationStatus.NotRun, requested));
        Assert.Equal("Invalid", model.FormatStandardsStatus(InvoiceValidationStatus.Invalid, requested));
        Assert.Equal("Passed", model.FormatStandardsStatus(InvoiceValidationStatus.Passed, requested));
    }

    private sealed class RecordingRunner : IOfficeInvoiceWorkflowRunner {
        internal OfficeInvoiceStorageWorkflowRequest? Request;
        internal CancellationToken Token;
        internal TaskCompletionSource Entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public async Task<OfficeInvoiceStorageWorkflowResult> RunInvoiceAsync(OfficeInvoiceStorageWorkflowRequest request,
            InvoiceValidator? validator = null, CancellationToken cancellationToken = default) {
            Request = request; Token = cancellationToken; Entered.TrySetResult();
            await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            throw new InvalidOperationException("The cancellation boundary should end execution.");
        }
    }
}
