using System.Reflection;
using System.Security.Cryptography;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPermissionStreamIdentityTests {
    [Theory]
    [InlineData("Snapshot", false)]
    [InlineData("Snapshot", true)]
    [InlineData("Identity", false)]
    [InlineData("Identity", true)]
    [InlineData("Publication", false)]
    [InlineData("Publication", true)]
    public async Task PermissionScopedFileReadsRetainOpenedIdentityAndRejectReplacement(string operation, bool replace) {
        using var root = new TestDirectory();
        string path = Path.Combine(root.Path, "source.pdf");
        byte[] original = StudioProviderDocumentTests.CreatePdf();
        byte[] replacement = StudioProviderDocumentTests.CreatePdf(2);
        File.WriteAllBytes(path, original);
        using var storage = new StudioStorageAccess();
        string originalIdentity = await storage.ReadIdentityAsync(path, default);
        FileStream? opened = null;
        Microsoft.Win32.SafeHandles.SafeFileHandle? handle = null;
        var file = new TestStorageFile(new Uri(path).AbsoluteUri, original) {
            OpenReadOverride = () => {
                opened = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                handle = opened.SafeFileHandle;
                Stream scoped = WrapActualPermissionStream(opened);
                if (replace) {
                    string other = Path.Combine(root.Path, "replacement.pdf");
                    File.WriteAllBytes(other, replacement);
                    File.Move(other, path, overwrite: true);
                }
                return Task.FromResult(scoped);
            }
        };
        string location = await storage.RegisterAsync(file.Item, default);
        async Task Run() {
            if (operation == "Snapshot") {
                StudioStorageSnapshot snapshot = await storage.ReadSnapshotAsync(location, default);
                Assert.Equal(original, snapshot.Bytes);
                Assert.Equal(originalIdentity, snapshot.Identity);
            } else if (operation == "Identity") {
                Assert.Equal(originalIdentity, await storage.ReadIdentityAsync(location, default));
            } else {
                string fingerprint = Convert.ToHexString(SHA256.HashData(original));
                StudioStoragePublication publication = await storage.PublishAsync(location, original, fingerprint,
                    _ => Task.CompletedTask, default);
                Assert.Equal(originalIdentity, publication.Identity);
            }
        }
        if (replace) {
            await Assert.ThrowsAsync<IOException>(Run);
            Assert.Equal(replacement, File.ReadAllBytes(path));
            Assert.Equal(0, file.Writes);
        } else {
            await Run();
        }
        Assert.True(handle!.IsClosed);
    }

    private static Stream WrapActualPermissionStream(FileStream file) {
        // Exercise the production wrapper without a test hook or a native grant.
        // A zero native URL is the permission owner's already-disposed state.
        Type wrapper = typeof(StudioStorageAccess).GetNestedType("PermissionStream", BindingFlags.NonPublic)!;
        ConstructorInfo constructor = wrapper.GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).Single();
        Type permissionType = constructor.GetParameters()[1].ParameterType;
        object permission = Activator.CreateInstance(permissionType, BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null, args: [IntPtr.Zero], culture: null)!;
        return (Stream)constructor.Invoke([file, permission]);
    }

    private sealed class TestDirectory : IDisposable {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "studio-permission-identity-" + Guid.NewGuid().ToString("N"));
        internal TestDirectory() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
