using OfficeIMO.Drawing.Internal;
using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class OfficeFileCommitTests {
        [Fact]
        public void Write_WhenProducerFails_PreservesExistingDestination() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".bin");
            byte[] original = { 1, 2, 3, 4 };
            File.WriteAllBytes(path, original);

            try {
                Assert.Throws<InvalidOperationException>(() => OfficeFileCommit.Write(path, stream => {
                    stream.WriteByte(9);
                    throw new InvalidOperationException("Simulated serialization failure.");
                }));

                Assert.Equal(original, File.ReadAllBytes(path));
                Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(path)!, $".{Path.GetFileName(path)}.*.tmp"));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Write_WithFailIfExists_DoesNotReplaceDestination() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".bin");
            byte[] original = { 1, 2, 3, 4 };
            File.WriteAllBytes(path, original);

            try {
                Assert.Throws<IOException>(() => OfficeFileCommit.WriteAllBytes(
                    path,
                    new byte[] { 9, 8, 7 },
                    OfficeFileCommit.ConflictPolicy.FailIfExists));

                Assert.Equal(original, File.ReadAllBytes(path));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void WriteAllBytes_CreatesMissingDestinationDirectory() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string path = Path.Combine(root, "nested", "artifact.bin");

            try {
                OfficeFileCommit.WriteAllBytes(path, new byte[] { 1, 2, 3, 4 });

                Assert.Equal(new byte[] { 1, 2, 3, 4 }, File.ReadAllBytes(path));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }

        [Fact]
        public void StagedBytesCanRetryAfterDestinationCollisionWithoutBeingRewritten() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string occupiedPath = Path.Combine(root, "artifact.bin");
            string availablePath = Path.Combine(root, "artifact-2.bin");
            byte[] payload = { 9, 8, 7, 6 };

            try {
                Directory.CreateDirectory(root);
                File.WriteAllBytes(occupiedPath, new byte[] { 1 });
                string stagingPath = OfficeFileCommit.StageAllBytes(occupiedPath, payload);

                Assert.False(OfficeFileCommit.TryCommitTemporaryFileIfAbsent(stagingPath, occupiedPath));
                Assert.True(File.Exists(stagingPath));
                Assert.Equal(payload, File.ReadAllBytes(stagingPath));

                Assert.True(OfficeFileCommit.TryCommitTemporaryFileIfAbsent(stagingPath, availablePath));
                Assert.False(File.Exists(stagingPath));
                Assert.Equal(payload, File.ReadAllBytes(availablePath));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }

        [Fact]
        public async Task WriteAllBytes_SyncAndAsync_PreserveReadOnlyDestinations() {
            string syncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".bin");
            string asyncPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".bin");
            byte[] original = { 1, 2, 3, 4 };
            File.WriteAllBytes(syncPath, original);
            File.WriteAllBytes(asyncPath, original);
            var syncDestination = new FileInfo(syncPath) { IsReadOnly = true };
            var asyncDestination = new FileInfo(asyncPath) { IsReadOnly = true };

            try {
                Assert.Throws<UnauthorizedAccessException>(() =>
                    OfficeFileCommit.WriteAllBytes(syncPath, new byte[] { 9, 8, 7 }));
                await Assert.ThrowsAsync<UnauthorizedAccessException>(() =>
                    OfficeFileCommit.WriteAllBytesAsync(asyncPath, new byte[] { 9, 8, 7 }));

                Assert.Equal(original, File.ReadAllBytes(syncPath));
                Assert.Equal(original, File.ReadAllBytes(asyncPath));
            } finally {
                syncDestination.IsReadOnly = false;
                asyncDestination.IsReadOnly = false;
                if (File.Exists(syncPath)) File.Delete(syncPath);
                if (File.Exists(asyncPath)) File.Delete(asyncPath);
            }
        }

        [Fact]
        public void CommitTemporaryFileAtomically_ReplacesAnExistingDestination() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string destination = Path.Combine(root, "artifact.bin");
            string temporary = Path.Combine(root, "artifact.staged.bin");
            Directory.CreateDirectory(root);
            File.WriteAllBytes(destination, new byte[] { 1, 2, 3, 4 });
            File.WriteAllBytes(temporary, new byte[] { 9, 8, 7 });

            try {
                OfficeFileCommit.CommitTemporaryFileAtomically(temporary, destination);

                Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(destination));
                Assert.False(File.Exists(temporary));
                Assert.Empty(Directory.GetFiles(root, "*.bak"));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }
    }
}
