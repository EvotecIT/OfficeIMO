using System;
using System.IO;
using OfficeIMO.Core.Internal;
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
    }
}
