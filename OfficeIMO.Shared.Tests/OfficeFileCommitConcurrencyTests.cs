using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public sealed class OfficeFileCommitConcurrencyTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task FailIfExistsPreservesTheOnlySuccessfulWritersBytes(bool asynchronous) {
            const int writers = 12;
            string root = Path.Combine(Path.GetTempPath(), "officeimo-commit-" + Guid.NewGuid().ToString("N"));
            try {
                for (int round = 0; round < 4; round++) {
                    string folder = Path.Combine(root, round.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    Directory.CreateDirectory(folder);
                    string path = Path.Combine(folder, "artifact.bin");
                    using var ready = new CountdownEvent(writers);
                    using var start = new ManualResetEventSlim(false);
                    Task<int?>[] attempts = Enumerable.Range(0, writers).Select(index =>
                        Task.Factory.StartNew(async () => {
                            byte[] payload = { (byte)index, 1, 2, 3 };
                            ready.Signal();
                            start.Wait();
                            try {
                                if (asynchronous) {
                                    await OfficeFileCommit.WriteAllBytesAsync(path, payload,
                                        OfficeFileCommit.ConflictPolicy.FailIfExists).ConfigureAwait(false);
                                } else {
                                    OfficeFileCommit.WriteAllBytes(path, payload,
                                        OfficeFileCommit.ConflictPolicy.FailIfExists);
                                }
                                return (int?)index;
                            } catch (IOException) {
                                return null;
                            }
                        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default).Unwrap())
                        .ToArray();

                    bool allWritersReady = ready.Wait(TimeSpan.FromSeconds(15));
                    start.Set();
                    int?[] results = await Task.WhenAll(attempts);
                    Assert.True(allWritersReady, "All file writers must reach the start barrier.");
                    int winner = Assert.Single(results, result => result.HasValue)!.Value;
                    Assert.Equal(new byte[] { (byte)winner, 1, 2, 3 }, File.ReadAllBytes(path));
                    Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(folder));
                    Directory.Delete(folder, recursive: true);
                }
            } finally {
                if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            }
        }
    }
}
