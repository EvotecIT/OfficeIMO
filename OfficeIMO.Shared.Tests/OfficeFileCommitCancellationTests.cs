using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class OfficeFileCommitCancellationTests {
        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public async Task CancellationAfterStagingPreservesDestinationAndRemovesTemporaryFiles(bool asynchronous, bool strictAtomic) {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-commit-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "existing.bin");
            byte[] original = { 1, 2, 3 };
            File.WriteAllBytes(path, original);
            using var cancellation = new CancellationTokenSource();
            try {
                void WriteAndCancel(Stream stream) {
                    stream.WriteByte(9);
                    cancellation.Cancel();
                }
                if (strictAtomic && asynchronous) {
                    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeFileCommit.WriteAtomicallyAsync(
                        path, (stream, token) => { WriteAndCancel(stream); return Task.CompletedTask; }, cancellation.Token));
                } else if (strictAtomic) {
                    Assert.ThrowsAny<OperationCanceledException>(() => OfficeFileCommit.WriteAtomically(path, WriteAndCancel, cancellation.Token));
                } else if (asynchronous) {
                    await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeFileCommit.WriteAsync(
                        path, WriteAndCancel, cancellationToken: cancellation.Token));
                } else {
                    Assert.ThrowsAny<OperationCanceledException>(() => OfficeFileCommit.Write(path, WriteAndCancel, cancellation.Token));
                }
                Assert.Equal(original, File.ReadAllBytes(path));
                Assert.Equal(new[] { path }, Directory.GetFiles(directory));
            } finally {
                Directory.Delete(directory, true);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task StrictPublicationCreatesAndReplacesCompleteFilesAndCleansUpProducerFailure(bool asynchronous) {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-atomic-" + Guid.NewGuid().ToString("N"));
            string path = Path.Combine(directory, "artifact.bin");
            try {
                async Task Publish(byte value, bool fail = false) {
                    void Produce(Stream stream) {
                        stream.WriteByte(value);
                        if (fail) throw new InvalidOperationException("Serialization failed after staging began.");
                    }
                    if (asynchronous) await OfficeFileCommit.WriteAtomicallyAsync(path,
                        (stream, token) => { Produce(stream); return Task.CompletedTask; });
                    else OfficeFileCommit.WriteAtomically(path, Produce);
                }

                await Publish(1);
                Assert.Equal(new byte[] { 1 }, File.ReadAllBytes(path));
                await Publish(2);
                Assert.Equal(new byte[] { 2 }, File.ReadAllBytes(path));
                await Assert.ThrowsAsync<InvalidOperationException>(() => Publish(3, fail: true));
                Assert.Equal(new byte[] { 2 }, File.ReadAllBytes(path));
                Assert.Equal(new[] { path }, Directory.GetFiles(directory));
            } finally {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }
    }
}
