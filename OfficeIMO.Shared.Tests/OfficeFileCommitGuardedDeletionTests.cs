using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public sealed class OfficeFileCommitGuardedDeletionTests {
        [Fact]
        public void MatchingSnapshotIsDeletedAndMissingDestinationSucceeds() {
            using var scope = new DeletionDirectory();
            File.WriteAllText(scope.Path, "Owned bytes");
            Assert.True(OfficeFileCommit.TryDeleteIfUnchanged(scope.Path, candidate => File.ReadAllText(candidate) == "Owned bytes"));
            Assert.Empty(Directory.GetFileSystemEntries(scope.Root));
            Assert.True(OfficeFileCommit.TryDeleteIfUnchanged(scope.Path, _ => throw new InvalidOperationException("A missing entry must not be validated.")));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ReplacementPublishedAfterSnapshotValidationSurvivesDeletion(bool lease) {
            using var scope = new DeletionDirectory(lease);
            File.WriteAllText(scope.Path, "Owned bytes");
            bool deleted = OfficeFileCommit.TryDeleteIfUnchanged(scope.Path, candidate => {
                bool matches = File.ReadAllText(candidate) == "Owned bytes";
                File.WriteAllText(scope.Path, "Another writer's replacement");
                return matches;
            });
            Assert.True(deleted);
            Assert.Equal("Another writer's replacement", File.ReadAllText(scope.Path));
            Assert.Equal(new[] { scope.Path }, Directory.GetFileSystemEntries(scope.Root));
        }

        [Fact]
        public void MismatchingSnapshotIsRestoredAndReturnsFalse() {
            using var scope = new DeletionDirectory();
            File.WriteAllText(scope.Path, "A changed report");
            Assert.False(OfficeFileCommit.TryDeleteIfUnchanged(scope.Path, candidate => File.ReadAllText(candidate) == "Earlier bytes"));
            Assert.Equal("A changed report", File.ReadAllText(scope.Path));
            Assert.Equal(new[] { scope.Path }, Directory.GetFileSystemEntries(scope.Root));
        }

        [Fact]
        public void ValidatorFailureRestoresSnapshotAndRethrowsOriginalError() {
            using var scope = new DeletionDirectory();
            File.WriteAllText(scope.Path, "Owned bytes");
            var original = new InvalidOperationException("Snapshot validation failed");
            var actual = Assert.Throws<InvalidOperationException>(() => OfficeFileCommit.TryDeleteIfUnchanged(scope.Path, _ => throw original));
            Assert.Same(original, actual);
            Assert.Equal("Owned bytes", File.ReadAllText(scope.Path));
            Assert.Equal(new[] { scope.Path }, Directory.GetFileSystemEntries(scope.Root));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void RestoreCollisionPreservesBothVersionsAndIdentifiesRetainedCandidate(bool validatorThrows) {
            using var scope = new DeletionDirectory();
            File.WriteAllText(scope.Path, "Displaced bytes");
            var error = Assert.Throws<IOException>(() => OfficeFileCommit.TryDeleteIfUnchanged(scope.Path, candidate => {
                Assert.Equal("Displaced bytes", File.ReadAllText(candidate));
                File.WriteAllText(scope.Path, "Newer destination bytes");
                if (validatorThrows) throw new InvalidOperationException("Validation failed after a newer save");
                return false;
            }));
            Assert.Equal("Newer destination bytes", File.ReadAllText(scope.Path));
            string retained = Assert.Single(Directory.GetFiles(scope.Root), path => path != scope.Path);
            Assert.Equal("Displaced bytes", File.ReadAllText(retained));
            Assert.Contains(retained, error.Message);
        }

        private sealed class DeletionDirectory : IDisposable {
            internal string Root { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-delete-" + Guid.NewGuid().ToString("N"));
            internal string Path { get; }
            internal DeletionDirectory(bool lease = false) {
                Directory.CreateDirectory(Root);
                Path = System.IO.Path.Combine(Root, lease ? "run.json.lease" : "run.json");
            }
            public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
        }
    }
}
