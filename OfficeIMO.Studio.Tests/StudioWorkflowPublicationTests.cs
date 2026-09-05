using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioWorkflowPublicationTests {
    [Fact]
    public async Task ConversionReplaceCannotOverwriteAnOpenTabThroughDirectoryAlias() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-publication-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "document.html");
            File.WriteAllText(input, "<html><body><p>Replacement</p></body></html>");
            string destinationFolder = Directory.CreateDirectory(Path.Combine(root, "documents")).FullName;
            string output = Path.Combine(destinationFolder, "document.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(output);
            byte[] original = File.ReadAllBytes(output);
            string alias = Path.Combine(root, "alias");
            Directory.CreateSymbolicLink(alias, destinationFolder);
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                StudioDocumentTabHost? host = null;
                var guard = new StudioWorkflowPublicationGuard((path, directory) => {
                    Assert.True(Dispatcher.UIThread.CheckAccess());
                    return directory ? host!.CanPublishDirectory(path) : host!.CanPublishPath(path);
                });
                host = new StudioDocumentTabHost(open => new MainWindowViewModel(
                    _ => Task.FromResult<string?>(null),
                    pickWorkflowFiles: _ => Task.FromResult<IReadOnlyList<string>>([input]),
                    openDocumentInTab: open, publicationGuard: guard), _ => { });
                using (host) {
                    await host.OpenDocumentAsync(output);
                    var tab = host.SelectedTab!;
                    tab.Document.SetOrganizerSelection([tab.Document.OrganizerPages[0]]);
                    await tab.Document.DuplicateSelectedCommand.ExecuteAsync(null);
                    var conversion = tab.Document.ConversionWorkbench;
                    conversion.SelectedRoute = conversion.Routes.Single(route => route.Route.Id == "html-pdf");
                    conversion.SelectedConflict = conversion.ConflictPolicies.Single(policy => policy.Value == OfficeWorkflowConflictPolicy.Replace);
                    conversion.OutputFolder = alias;
                    await conversion.AddFilesCommand.ExecuteAsync(null);
                    await conversion.RunQueueCommand.ExecuteAsync(null);
                    Assert.Equal("Failed", Assert.Single(conversion.Jobs).Status);
                    Assert.Equal(original, File.ReadAllBytes(output));
                    Assert.True(tab.Document.IsDirty);
                    Assert.Equal(2, tab.Document.Pages.Count);
                    Assert.False(host.CanPublishDirectory(alias));
                    Assert.False(host.CanPublishDirectory(root));
                    Assert.True(host.CanPublishDirectory(Path.Combine(root, "other")));

                    conversion.SelectedConflict = conversion.ConflictPolicies.Single(policy => policy.Value == OfficeWorkflowConflictPolicy.Rename);
                    await conversion.RetryFailedCommand.ExecuteAsync(null);
                    Assert.Equal("Completed", Assert.Single(conversion.Jobs).Status);
                    Assert.Equal(original, File.ReadAllBytes(output));
                    Assert.Equal(1, PdfDocument.Load(File.ReadAllBytes(Path.Combine(destinationFolder, "document (1).pdf"))).Inspect().PageCount);
                }
                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task WindowCompositionChecksLiveOwnershipForWorkflowPublication() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-window-publication-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string output = Path.Combine(root, "assembled.pdf");
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(source);
            File.Copy(source, output);
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var window = new MainWindow(TestAppBuilder.CreateTestServices());
                try {
                    await window.TabHost.OpenDocumentAsync(output);
                    // A live tab still owns its pathname after external deletion.
                    File.Delete(output);
                    var assembly = window.ViewModel.OutputWorkbench.Assembly;
                    assembly.UseDocument(source);
                    assembly.OutputPath = output;
                    await assembly.RunCommand.ExecuteAsync(null);
                    Assert.Equal(Path.Combine(root, "assembled (1).pdf"), assembly.PublishedPath);
                    Assert.False(File.Exists(output));
                    Assert.Single(window.TabHost.Tabs);
                } finally {
                    window.Close();
                    window.TabHost.Dispose();
                }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
