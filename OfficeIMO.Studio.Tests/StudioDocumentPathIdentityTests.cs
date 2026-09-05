using System.Runtime.InteropServices;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioDocumentPathIdentityTests {
    [Theory]
    [InlineData("hard-link")]
    [InlineData("file-link")]
    [InlineData("directory-link")]
    public async Task AliasesReuseLiveEditsAndCannotOverwriteAnotherTab(string kind) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        StudioDocumentTabHost? host = null;
        try {
            string sourceDirectory = Directory.CreateDirectory(Path.Combine(root, "source")).FullName;
            string source = Path.Combine(sourceDirectory, "document.pdf");
            string other = Path.Combine(root, "other.pdf");
            CreateDocument(source);
            CreateDocument(other);
            byte[] original = File.ReadAllBytes(source);
            string alias = Path.Combine(root, "alias.pdf");
            if (kind == "hard-link") {
                bool created = OperatingSystem.IsWindows()
                    ? CreateHardLink(alias, source, IntPtr.Zero)
                    : Link(source, alias) == 0;
                Assert.True(created, $"Hard link creation failed: {Marshal.GetLastPInvokeError()}");
            } else if (kind == "file-link") {
                File.CreateSymbolicLink(alias, source);
            } else {
                string directoryAlias = Path.Combine(root, "alias-directory");
                Directory.CreateSymbolicLink(directoryAlias, sourceDirectory);
                alias = Path.Combine(directoryAlias, "document.pdf");
            }
            host = new StudioDocumentTabHost(open => {
                MainWindowViewModel? document = null;
                document = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                    pickSavePdf: _ => Task.FromResult<string?>(alias),
                    canSaveAsPath: path => host!.CanDocumentOwnPath(document, path),
                    openDocumentInTab: open);
                return document;
            }, _ => { });
            await host.OpenDocumentAsync(source);
            var sourceTab = Assert.Single(host.Tabs);
            sourceTab.Document.SetOrganizerSelection([sourceTab.Document.OrganizerPages[0]]);
            await sourceTab.Document.DuplicateSelectedCommand.ExecuteAsync(null);
            await host.OpenDocumentAsync(alias);
            Assert.Same(sourceTab, Assert.Single(host.Tabs));
            Assert.True(sourceTab.Document.IsDirty);
            Assert.Equal(2, sourceTab.Document.Pages.Count);
            Assert.False(host.CanPublishPath(alias));
            Assert.True(host.CanDocumentOwnPath(sourceTab.Document, alias));

            await host.OpenDocumentAsync(other);
            var otherTab = host.SelectedTab!;
            Assert.False(host.CanDocumentOwnPath(otherTab.Document, alias));
            await otherTab.Document.SaveAsCommand.ExecuteAsync(null);
            Assert.Equal(other, otherTab.Document.DocumentPath);
            Assert.Contains("already open", otherTab.Document.OperationStatus, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(original, File.ReadAllBytes(alias));
        } finally {
            host?.Dispose();
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task UnresolvedLinkPreservesExistingDirtyTabAndRejectsPublication() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-unresolved-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "document.pdf");
            string cyclic = Path.Combine(root, "cyclic.pdf");
            CreateDocument(source);
            File.CreateSymbolicLink(cyclic, cyclic);
            using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
                _ => Task.FromResult<string?>(null), openDocumentInTab: open), _ => { });
            await host.OpenDocumentAsync(source);
            var tab = Assert.Single(host.Tabs);
            tab.Document.SetOrganizerSelection([tab.Document.OrganizerPages[0]]);
            await tab.Document.DuplicateSelectedCommand.ExecuteAsync(null);

            Assert.False(host.CanPublishPath(cyclic));
            await host.OpenDocumentAsync(cyclic);

            Assert.Same(tab, Assert.Single(host.Tabs));
            Assert.Same(tab, host.SelectedTab);
            Assert.True(tab.Document.IsDirty);
            Assert.Equal(2, tab.Document.Pages.Count);
            Assert.False(string.IsNullOrWhiteSpace(tab.Document.ErrorMessage));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task InvalidPathsDoNotAuthorizePublicationOrDisruptAnOpenTab() {
        using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
            _ => Task.FromResult<string?>(null), openDocumentInTab: open), _ => { });
        Assert.False(host.CanPublishPath("invalid\0.pdf"));
        await host.OpenDocumentAsync("invalid\0.pdf");
        Assert.Empty(host.Tabs);
        Assert.False(string.IsNullOrWhiteSpace(host.ActiveDocument.ErrorMessage));
    }

    private static void CreateDocument(string path) =>
        PdfDocument.Create(document => document.Page(page => page.Size(420D, 620D))).Save(path);

    [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string newFile, string existingFile, IntPtr securityAttributes);

    [DllImport("libc", EntryPoint = "link", SetLastError = true)]
    private static extern int Link(string existingFile, string newFile);
}
