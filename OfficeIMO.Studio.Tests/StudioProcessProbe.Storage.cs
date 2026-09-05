using OfficeIMO.Pdf;
using System.Diagnostics;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

internal static partial class StudioProcessProbe {
    private static async Task VerifyDetachedStorageAsync(string root, StudioApplicationServices services) {
        string volume = Path.Combine(root, "volume");
        Assert.True(OperatingSystem.IsLinux());
        Assert.Equal("tmpfs", new DriveInfo(volume).DriveFormat);
        Assert.InRange(new DriveInfo(volume).TotalSize, 1, 32L * 1024 * 1024);
        string? parentNamespace = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_PARENT_MOUNT_NAMESPACE");
        Assert.False(string.IsNullOrWhiteSpace(parentNamespace));
        Assert.NotEqual(parentNamespace, new FileInfo("/proc/self/ns/mnt").LinkTarget);
        string source = Path.Combine(volume, "source.pdf");
        string destination = Path.Combine(root, "recovered.pdf");
        PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(source);
        byte[] original = File.ReadAllBytes(source);
        using (var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
            _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services), _ => { })) {
            using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
            await host.OpenDocumentAsync(source);
            await DuplicateAsync(host.ActiveDocument);
            session.Flush();
        }
        var info = new ProcessStartInfo("umount") { UseShellExecute = false, RedirectStandardError = true };
        info.ArgumentList.Add(volume);
        using (var unmount = Process.Start(info) ?? throw new IOException("Could not detach the acceptance volume.")) {
            Task<string> errors = unmount.StandardError.ReadToEndAsync();
            try {
                await unmount.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                Assert.True(unmount.ExitCode == 0, await errors);
            } finally {
                if (!unmount.HasExited) { unmount.Kill(); await unmount.WaitForExitAsync(); }
            }
        }
        Assert.False(File.Exists(source));
        // The backing mount lets us verify that detaching the document path did not change its bytes.
        Assert.Equal(original, File.ReadAllBytes(Path.Combine(root, "backing", "source.pdf")));
        using var restarted = new StudioDocumentTabHost(open => new MainWindowViewModel(
            _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services), _ => { });
        using var restart = new StudioSessionController(restarted, services, _ => Task.FromResult<string?>(destination));
        await restart.InspectAsync();
        var item = Assert.Single(restart.Pending);
        Assert.False(item.SourceExists);
        Assert.False(item.SourceUnchanged);
        Assert.True(item.HasRecovery);
        await restart.RecoverCopyCommand.ExecuteAsync(item);
        Assert.Empty(restart.Pending);
        Assert.False(restart.HasError);
        Assert.Equal(2, PdfDocument.Load(destination).Read().Pages.Count);
        Assert.Equal(original, File.ReadAllBytes(Path.Combine(root, "backing", "source.pdf")));
    }

    private static async Task VerifyFullStorageAsync(string root, StudioApplicationServices services) {
        string volume = Path.Combine(root, "volume");
        var drive = new DriveInfo(volume);
        // Never exhaust an ordinary filesystem: this probe accepts only a bounded disposable tmpfs.
        Assert.True(OperatingSystem.IsLinux());
        Assert.Equal("tmpfs", drive.DriveFormat);
        Assert.InRange(drive.TotalSize, 1, 32L * 1024 * 1024);
        string source = Path.Combine(root, "source.pdf");
        string destination = Path.Combine(volume, "saved.pdf");
        string filler = Path.Combine(volume, "fill.bin");
        PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(source);
        byte[] original = File.ReadAllBytes(source);
        File.WriteAllBytes(destination, original);
        using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
            _ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(destination),
            openDocumentInTab: open, services: services), _ => { });
        using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
        await host.OpenDocumentAsync(source);
        await DuplicateAsync(host.ActiveDocument);
        session.Flush();
        byte[] previousSession = File.ReadAllBytes(services.Paths.SessionPath);
        string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(original);
        byte[] previousRecovery = Assert.IsType<byte[]>(services.Recovery.ReadVerifiedSnapshot(source, fingerprint));
        FillVolume(filler, drive);
        try {
            var document = host.ActiveDocument;
            document.SetOrganizerSelection([document.OrganizerPages[0]]);
            await document.DuplicateSelectedCommand.ExecuteAsync(null);
            Assert.True(document.HasError);
            Assert.Equal(2, document.Pages.Count);
            Assert.True(document.IsDirty);
            Assert.Equal(previousRecovery, services.Recovery.ReadVerifiedSnapshot(source, fingerprint));
            Assert.Equal(original, File.ReadAllBytes(source));
            await document.SaveAsCommand.ExecuteAsync(null);
            Assert.True(document.HasError);
            Assert.True(document.IsDirty);
            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Equal(previousRecovery, services.Recovery.ReadVerifiedSnapshot(source, fingerprint));
            session.Flush();
            Assert.Equal(previousSession, File.ReadAllBytes(services.Paths.SessionPath));
            Assert.True(session.HasStorageError, "A failed session write must be visible to the user.");
            Assert.Empty(Directory.EnumerateFiles(volume, "*.tmp", SearchOption.AllDirectories));
        } finally { File.Delete(filler); }
        session.Flush();
        Assert.False(session.HasStorageError);
        await host.ActiveDocument.SaveAsCommand.ExecuteAsync(null);
        Assert.False(host.ActiveDocument.HasError);
        Assert.False(host.ActiveDocument.IsDirty);
        Assert.Equal(2, PdfDocument.Load(destination).Read().Pages.Count);
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    private static void FillVolume(string filler, DriveInfo drive) {
        using var stream = new FileStream(filler, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        byte[] block = new byte[64 * 1024];
        bool exhausted = false;
        try {
            for (int index = 0; index < 1024; index++) stream.Write(block);
        } catch (IOException) { exhausted = true; }
        Assert.True(exhausted, "The bounded filesystem did not report a write failure.");
        Assert.InRange(drive.AvailableFreeSpace, 0, 4095);
    }
}
