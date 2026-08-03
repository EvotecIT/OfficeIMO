using OfficeIMO.Drawing.Internal;
using System;
using System.IO;
using System.Linq;
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

        [Fact]
        public void GuardedAtomicCommit_RestoresAChangedDestination() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string destination = Path.Combine(root, "artifact.bin");
            string temporary = Path.Combine(root, "artifact.staged.bin");
            byte[] expectedOriginal = { 1, 2, 3, 4 };
            byte[] concurrentEdit = { 5, 6, 7, 8 };
            Directory.CreateDirectory(root);
            File.WriteAllBytes(destination, concurrentEdit);
            File.WriteAllBytes(temporary, new byte[] { 9, 8, 7 });

            try {
                bool committed = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    temporary,
                    destination,
                    displaced => File.ReadAllBytes(displaced).SequenceEqual(expectedOriginal));

                Assert.False(committed);
                Assert.Equal(concurrentEdit, File.ReadAllBytes(destination));
                Assert.False(File.Exists(temporary));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }

        [Fact]
        public void GuardedAtomicCommit_RestoresDestinationWhenInstalledStageDoesNotMatch() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string destination = Path.Combine(root, "artifact.bin");
            string temporary = Path.Combine(root, "artifact.staged.bin");
            byte[] expectedOriginal = { 1, 2, 3, 4 };
            byte[] expectedStage = { 9, 8, 7 };
            byte[] replacedStage = { 6, 6, 6 };
            Directory.CreateDirectory(root);
            File.WriteAllBytes(destination, expectedOriginal);
            File.WriteAllBytes(temporary, replacedStage);

            try {
                bool committed = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    temporary,
                    destination,
                    displaced => File.ReadAllBytes(displaced).SequenceEqual(expectedOriginal),
                    installed => File.ReadAllBytes(installed).SequenceEqual(expectedStage));

                Assert.False(committed);
                Assert.Equal(expectedOriginal, File.ReadAllBytes(destination));
                Assert.False(File.Exists(temporary));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }

        [Fact]
        public void GuardedAtomicCommit_PreservesSaveThatRacesWithRollback() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string destination = Path.Combine(root, "artifact.bin");
            string temporary = Path.Combine(root, "artifact.staged.bin");
            byte[] expectedOriginal = { 1, 2, 3, 4 };
            byte[] staged = { 9, 8, 7 };
            byte[] newerSave = { 5, 6, 7, 8 };
            Directory.CreateDirectory(root);
            File.WriteAllBytes(destination, expectedOriginal);
            File.WriteAllBytes(temporary, staged);

            try {
                bool committed = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    temporary,
                    destination,
                    displaced => {
                        Assert.Equal(expectedOriginal, File.ReadAllBytes(displaced));
                        File.WriteAllBytes(destination, newerSave);
                        return false;
                    });

                Assert.False(committed);
                Assert.Equal(newerSave, File.ReadAllBytes(destination));
                Assert.False(File.Exists(temporary));
                Assert.Single(Directory.GetFiles(root));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }

        [Fact]
        public void GuardedAtomicCommit_PreservesSaveDisplacedBySecondRollbackReplacement() {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string destination = Path.Combine(root, "artifact.bin");
            string temporary = Path.Combine(root, "artifact.staged.bin");
            byte[] expectedOriginal = { 1, 2, 3, 4 };
            byte[] staged = { 9, 8, 7 };
            byte[] firstConcurrentSave = { 5, 6, 7, 8 };
            byte[] secondConcurrentSave = { 4, 3, 2, 1 };
            Directory.CreateDirectory(root);
            File.WriteAllBytes(destination, expectedOriginal);
            File.WriteAllBytes(temporary, staged);

            try {
                IOException exception = Assert.Throws<IOException>(() =>
                    OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchangedForTesting(
                        temporary,
                        destination,
                        displaced => {
                            Assert.Equal(expectedOriginal, File.ReadAllBytes(displaced));
                            File.WriteAllBytes(destination, firstConcurrentSave);
                            return false;
                        },
                        installedFileMatchesExpected: null,
                        afterFirstRollbackReplacement: target => File.WriteAllBytes(target, secondConcurrentSave)));

                Assert.Equal(firstConcurrentSave, File.ReadAllBytes(destination));
                string preservedPath = Assert.Single(
                    Directory.GetFiles(root),
                    path => !string.Equals(path, destination, StringComparison.Ordinal));
                Assert.Equal(secondConcurrentSave, File.ReadAllBytes(preservedPath));
                Assert.Contains(preservedPath, exception.Message, StringComparison.Ordinal);
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }

#if NET6_0_OR_GREATER
        [Fact]
        public void CommitTemporaryFile_PreservesRestrictiveUnixMode() {
            if (OperatingSystem.IsWindows()) return;

            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string destination = Path.Combine(root, "private.xlsx");
            string temporary = Path.Combine(root, "private.staged.xlsx");
            Directory.CreateDirectory(root);
            File.WriteAllBytes(destination, new byte[] { 1 });
            File.SetUnixFileMode(destination, UnixFileMode.UserRead | UnixFileMode.UserWrite);
            File.WriteAllBytes(temporary, new byte[] { 2 });

            try {
                OfficeFileCommit.CommitTemporaryFile(temporary, destination);

                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(destination));
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }
        }
#endif
    }
}
