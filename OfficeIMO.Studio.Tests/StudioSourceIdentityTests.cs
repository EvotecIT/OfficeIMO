using System.Runtime.InteropServices;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSourceIdentityTests {
    [Theory]
    [InlineData("save-as")]
    [InlineData("extract")]
    [InlineData("protect")]
    [InlineData("decrypt")]
    [InlineData("split")]
    public async Task WorkspaceOutputsRecheckPublicationOwnershipAfterStaging(string operation) {
        string root = NewFolder();
        try {
            string source = Path.Combine(root, "source.pdf");
            CreatePdf(source);
            File.WriteAllBytes(source, PdfDocument.Load(File.ReadAllBytes(source)).Security.Encrypt(
                new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }).Pdf);
            byte[] original = File.ReadAllBytes(source);
            string outputFolder = Directory.CreateDirectory(Path.Combine(root, "output")).FullName;
            string destination = Path.Combine(outputFolder, "copy.pdf");
            if (operation != "split") File.WriteAllText(destination, "Existing output retained");
            int checks = 0;
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None,
                new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")), "owner",
                (_, _) => ValueTask.FromResult(++checks == 1));
            if (operation != "decrypt") await workspace.DuplicateAsync([1], CancellationToken.None);
            if (operation == "split") {
                var result = await workspace.SplitAsync(outputFolder, 1, CancellationToken.None);
                Assert.False(result.Succeeded);
                Assert.Empty(result.Files);
            } else await Assert.ThrowsAsync<IOException>(async () => {
                switch (operation) {
                    case "save-as": await workspace.SaveAsync(destination, CancellationToken.None); break;
                    case "extract": await workspace.ExtractAsync([1], destination, CancellationToken.None); break;
                    case "protect": await workspace.SaveProtectedCopyAsync(destination, new PdfStandardEncryptionOptions("new"), "owner", CancellationToken.None); break;
                    case "decrypt": await workspace.SaveDecryptedCopyAsync(destination, "owner", CancellationToken.None); break;
                }
            });
            Assert.Equal(2, checks);
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(operation != "decrypt", workspace.IsDirty);
            Assert.Equal(source, workspace.Path);
            if (operation == "split") Assert.Empty(Directory.GetFileSystemEntries(outputFolder));
            else {
                Assert.Equal("Existing output retained", File.ReadAllText(destination));
                Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(outputFolder));
            }
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("changed")]
    [InlineData("replaced")]
    [InlineData("missing")]
    public async Task SaveRejectsExternalChangesAndSaveAsRetainsBothVersions(string change) {
        string root = NewFolder();
        try {
            string path = Path.Combine(root, "source.pdf");
            CreatePdf(path);
            byte[] original = File.ReadAllBytes(path);
            using var workspace = await PdfWorkspace.OpenAsync(path, CancellationToken.None, new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));
            await workspace.DuplicateAsync([1], CancellationToken.None);
            byte[] edits = workspace.CopyBytes();
            byte[] external = PdfDocument.Create(document => document.Page(page => page.Size(250, 350))).ToBytes();
            if (change == "changed") File.WriteAllBytes(path, external);
            else {
                File.Move(path, Path.Combine(root, "old.pdf"));
                if (change == "replaced") File.WriteAllBytes(path, original);
            }
            await Assert.ThrowsAnyAsync<IOException>(() => workspace.SaveAsync(null, CancellationToken.None));
            Assert.True(workspace.IsDirty);
            Assert.Equal(edits, workspace.CopyBytes());
            if (change == "changed") Assert.Equal(external, File.ReadAllBytes(path));
            else if (change == "replaced") Assert.Equal(original, File.ReadAllBytes(path));
            else Assert.False(File.Exists(path));
            string copy = Path.Combine(root, "my-edits.pdf");
            await workspace.SaveAsync(copy, CancellationToken.None);
            Assert.False(workspace.IsDirty);
            Assert.Equal(2, PdfDocument.Load(File.ReadAllBytes(copy)).Inspect().PageCount);
            await workspace.SaveAsync(null, CancellationToken.None);
            Assert.False(workspace.IsDirty);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task SavingThroughASymlinkPreservesTheLinkAndUpdatesTheSource() {
        string root = NewFolder();
        try {
            string source = Path.Combine(root, "source.pdf"), alias = Path.Combine(root, "alias.pdf");
            CreatePdf(source);
            File.CreateSymbolicLink(alias, source);
            using var workspace = await PdfWorkspace.OpenAsync(alias, CancellationToken.None, new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.SaveAsync(null, CancellationToken.None);
            Assert.NotNull(new FileInfo(alias).LinkTarget);
            Assert.Equal(2, PdfDocument.Load(File.ReadAllBytes(source)).Inspect().PageCount);
            Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(alias));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExtractionAndSecurityCopiesRejectAliasesOfTheirSource(bool hardLink) {
        string root = NewFolder();
        try {
            string source = Path.Combine(root, "source.pdf"), alias = Path.Combine(root, "alias.pdf");
            CreatePdf(source);
            byte[] plain = File.ReadAllBytes(source);
            byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }).Pdf;
            File.WriteAllBytes(source, encrypted);
            Alias(source, alias, hardLink);
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None,
                new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")), "owner");
            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.ExtractAsync([1], alias, CancellationToken.None));
            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.SaveProtectedCopyAsync(alias,
                new PdfStandardEncryptionOptions("new"), "owner", CancellationToken.None));
            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.SaveDecryptedCopyAsync(alias, "owner", CancellationToken.None));
            Assert.Equal(encrypted, File.ReadAllBytes(source));
            Assert.Equal(encrypted, File.ReadAllBytes(alias));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task AssemblyDeduplicatesPhysicalSourcesAcrossPickerCallsAndPreservesCaseDistinctInputs() {
        string root = NewFolder();
        try {
            string source = Path.Combine(root, "source.pdf"), alias = Path.Combine(root, "alias.pdf"), hard = Path.Combine(root, "hard.pdf");
            CreatePdf(source);
            Alias(source, alias, false); Alias(source, hard, true);
            string upper = Path.Combine(root, "SOURCE.pdf");
            bool sameCaseFile = File.Exists(upper);
            if (!sameCaseFile) CreatePdf(upper);
            IReadOnlyList<string> picked = [source, alias];
            using var assembly = new PdfAssemblyViewModel(_ => Task.FromResult(picked), _ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null));
            await assembly.AddFilesCommand.ExecuteAsync(null);
            Assert.Single(assembly.Sources);
            picked = [hard, upper, source];
            await assembly.AddFilesCommand.ExecuteAsync(null);
            Assert.Equal(sameCaseFile ? 1 : 2, assembly.Sources.Count);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ComparisonDoesNotOpenAnAliasOfTheCurrentDocument() {
        string root = NewFolder();
        try {
            string source = Path.Combine(root, "source.pdf"), alias = Path.Combine(root, "alias.pdf");
            CreatePdf(source); Alias(source, alias, true);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(alias), services: TestAppBuilder.CreateTestServices());
            await model.OpenDocumentAsync(source);
            await model.OpenComparisonCommand.ExecuteAsync(null);
            Assert.False(model.IsComparisonOpen);
            Assert.Empty(model.ComparisonPages);
            Assert.NotNull(model.OperationStatus);
        } finally { Directory.Delete(root, true); }
    }

    private static string NewFolder() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-source-identity-" + Guid.NewGuid().ToString("N"))).FullName;
    private static void CreatePdf(string path) => PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(path);
    private static void Alias(string source, string alias, bool hard) {
        if (!hard) File.CreateSymbolicLink(alias, source);
        else if (OperatingSystem.IsWindows()) Assert.True(CreateHardLink(alias, source, IntPtr.Zero));
        else Assert.Equal(0, Link(source, alias));
    }
    [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string newFile, string existingFile, IntPtr securityAttributes);
    [DllImport("libc", EntryPoint = "link", SetLastError = true)]
    private static extern int Link(string existingFile, string newFile);
}
