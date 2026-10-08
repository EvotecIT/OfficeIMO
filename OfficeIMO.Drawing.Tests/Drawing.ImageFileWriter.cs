using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageFileWriterTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task PublicWriterCreatesAndReplacesCompleteBorrowedBytes(bool asynchronous, bool existing) {
        await WithImageDirectory(async directory => {
            string path = Path.Combine(directory, "output.png");
            if (existing) {
                File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            }
            byte[] bytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 2), OfficeImageExportFormat.Png);
            byte[] original = (byte[])bytes.Clone();
            await Save(path, bytes, asynchronous);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(original, bytes);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicWriterPreCancellationPreservesDestinationAndCreatesNoStaging(bool asynchronous) {
        await WithImageDirectory(async directory => {
            string path = Path.Combine(directory, "output.png");
            byte[] original = { 1, 2, 3 };
            File.WriteAllBytes(path, original);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Save(path, new byte[] { 4, 5, 6 }, asynchronous, cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicWriterFailedDestinationCommitRemovesStaging(bool asynchronous) {
        await WithImageDirectory(async directory => {
            string path = Path.Combine(directory, "directory.png");
            Directory.CreateDirectory(path);
            Exception? failure = await Record.ExceptionAsync(() => Save(path, new byte[] { 4, 5, 6 }, asynchronous));
            Assert.True(failure is IOException || failure is UnauthorizedAccessException,
                "A directory destination must fail as a filesystem operation.");
            Assert.True(Directory.Exists(path));
            Assert.Empty(Directory.GetFiles(directory));
            Assert.Empty(Directory.GetFiles(path));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicWriterLockedDestinationPreservesBytesAndRemovesStaging(bool asynchronous) {
        await WithImageDirectory(async directory => {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return;
            }
            string path = Path.Combine(directory, "output.png");
            byte[] original = { 1, 2, 3 };
            File.WriteAllBytes(path, original);
            using (var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                await Assert.ThrowsAnyAsync<IOException>(() => Save(path, new byte[] { 4, 5, 6 }, asynchronous));
            }
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        });
    }

    private static Task Save(string path, byte[] bytes, bool asynchronous, CancellationToken token = default) {
        if (asynchronous) {
            return OfficeImageFileWriter.WriteAllBytesAsync(path, bytes, token);
        }
        OfficeImageFileWriter.WriteAllBytes(path, bytes, token);
        return Task.CompletedTask;
    }

    private static async Task WithImageDirectory(Func<string, Task> operation) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-image-file-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            await operation(directory);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
