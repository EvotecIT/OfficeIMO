using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioWorkspaceOutputOwnershipTests {
    [Theory]
    [InlineData("save-as")]
    [InlineData("extract")]
    [InlineData("protect")]
    [InlineData("split")]
    public async Task WorkspaceCommandsCannotPublishOverAnotherTabsPendingEdits(string operation) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            string destination = operation == "split" ? Path.Combine(services.Paths.Root, "Split PDFs", "part-001.pdf")
                : Path.Combine(services.Paths.Root, "owned.pdf");
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(source);
            File.Copy(source, destination);
            byte[] original = File.ReadAllBytes(destination);
            StudioDocumentTabHost? host = null;
            var guard = new StudioWorkflowPublicationGuard((path, directory) => directory ? host!.CanPublishDirectory(path) : host!.CanPublishPath(path));
            host = new StudioDocumentTabHost(open => new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                services: services, openDocumentInTab: open, publicationGuard: guard,
                pickSavePdf: _ => Task.FromResult<string?>(destination),
                pickOutputFolder: _ => Task.FromResult<string?>(services.Paths.Root), reviewPageSplit: _ => Task.FromResult(true)), _ => { });
            using (host) {
                await host.OpenDocumentAsync(source);
                var producer = host.ActiveDocument;
                producer.SetOrganizerSelection([producer.OrganizerPages[0]]);
                await host.OpenDocumentAsync(destination);
                var other = host.ActiveDocument;
                other.SetOrganizerSelection([other.OrganizerPages[0]]);
                await other.DuplicateSelectedCommand.ExecuteAsync(null);
                if (operation == "split") File.Delete(destination);
                switch (operation) {
                    case "save-as": await producer.SaveAsCommand.ExecuteAsync(null); break;
                    case "extract": await producer.ExtractSelectedCommand.ExecuteAsync(null); break;
                    case "protect":
                        producer.ProtectUserPassword = "new"; producer.ProtectConfirmPassword = "new";
                        await producer.SaveProtectedCopyCommand.ExecuteAsync(null); break;
                    case "split": await producer.SplitCommand.ExecuteAsync(null); break;
                }
                if (operation == "split") {
                    Assert.Null(producer.ErrorMessage);
                    Assert.True(Assert.Single(services.Jobs.Entries).HasOutput);
                    Assert.NotEqual(Path.GetDirectoryName(destination), Assert.Single(services.Jobs.Entries).OutputPath);
                } else Assert.Contains("open document", producer.ErrorMessage);
                Assert.True(other.IsDirty);
                Assert.Equal(2, other.Pages.Count);
                Assert.Equal(2, host.Tabs.Count);
                Assert.Equal(source, producer.DocumentPath);
                if (operation == "split") Assert.False(File.Exists(destination));
                else Assert.Equal(original, File.ReadAllBytes(destination));
            }
            return true;
        }, CancellationToken.None);
    }
}
