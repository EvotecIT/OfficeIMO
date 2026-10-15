using System.Text;
using System.Security.Cryptography;
using Avalonia.Threading;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioLocalPublicationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task VerifiedPublicationPreservesAConcurrentDestinationAndRemovesStaging(bool existing) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string directory = Path.Combine(((App)Avalonia.Application.Current!).Services.Paths.Root, "verified-publication");
            Directory.CreateDirectory(directory);
            string output = Path.Combine(directory, "book.oibook");
            byte[] original = Encoding.UTF8.GetBytes("Original destination");
            if (existing) File.WriteAllBytes(output, original);
            string? expected = existing ? Convert.ToHexString(SHA256.HashData(original)) : null;
            int authorizations = 0;
            await Assert.ThrowsAsync<IOException>(() => Task.Run(() => StudioLocalPublication.WriteVerifiedAsync(output,
                Encoding.UTF8.GetBytes("Reviewed output"), expected, async token => {
                    await Task.Yield();
                    Assert.True(Dispatcher.UIThread.CheckAccess());
                    token.ThrowIfCancellationRequested();
                    if (++authorizations == 2) File.WriteAllText(output, "Concurrent destination");
                }, CancellationToken.None)));
            Assert.Equal("Concurrent destination", File.ReadAllText(output));
            Assert.Single(Directory.GetFiles(directory));
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RejectedPublicationPreservesDestinationAndRemovesStaging(bool existing) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-assistant-publication-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string output = Path.Combine(directory, "answer.txt");
            try {
                if (existing) File.WriteAllText(output, "Keep this answer");
                await Assert.ThrowsAsync<IOException>(() => StudioLocalPublication.WriteAsync(output, Encoding.UTF8.GetBytes("New answer"), () => {
                    Assert.True(Dispatcher.UIThread.CheckAccess());
                    string staged = Assert.Single(Directory.GetFiles(directory), path => path != output);
                    Assert.Equal("New answer", File.ReadAllText(staged));
                    throw new IOException("The source or destination changed during staging.");
                }, CancellationToken.None));
                if (existing) Assert.Equal("Keep this answer", File.ReadAllText(output));
                else Assert.False(File.Exists(output));
                Assert.Equal(existing ? 1 : 0, Directory.GetFiles(directory).Length);
            } finally { Directory.Delete(directory, recursive: true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CurrentPublicationCommitsFlushedAnswerFromBackgroundCaller() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-assistant-publication-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string output = Path.Combine(directory, "answer.txt");
            try {
                File.WriteAllText(output, "Previous answer");
                await Task.Run(() => StudioLocalPublication.WriteAsync(output, Encoding.UTF8.GetBytes("Source hash and cited answer"),
                    () => Assert.True(Dispatcher.UIThread.CheckAccess()), CancellationToken.None));
                Assert.Equal("Source hash and cited answer", File.ReadAllText(output));
                Assert.Single(Directory.GetFiles(directory));
            } finally { Directory.Delete(directory, recursive: true); }
            return true;
        }, CancellationToken.None);
    }
}
