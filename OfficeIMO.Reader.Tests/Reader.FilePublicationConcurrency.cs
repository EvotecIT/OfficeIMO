using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Reader;
using Xunit;
using CoreCommitClaim = OfficeIMO.Core.Internal.IO.OfficeFileCommitClaim;
using ReaderCommitClaim = OfficeIMO.Reader.Internal.IO.OfficeFileCommitClaim;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderFilePublicationConcurrencyTests {
    [Theory]
    [InlineData("asset", ".bin")]
    [InlineData("table", ".csv")]
    [InlineData("visual", ".mmd")]
    public Task PublicMaterializersPreserveTheWinningPayloadAcrossConcurrentWriters(string materializer, string extension) =>
        AssertSinglePublisherAsync(extension, (folder, writer) => CreateReaderPublisher(materializer, folder, writer));

    [Fact]
    public Task CoreAndReaderPublishersShareNoClobberExclusion() {
        Assert.NotSame(typeof(OfficeFileCommit).Assembly, typeof(OfficeDocumentAssetMaterializer).Assembly);
        return AssertSinglePublisherAsync(".bin", (folder, writer) => {
            if (writer % 2 != 0) return CreateReaderPublisher("asset", folder, writer);

            byte[] payload = Encoding.UTF8.GetBytes("Core writer " + writer.ToString(System.Globalization.CultureInfo.InvariantCulture));
            string path = Path.Combine(folder, "shared-export.bin");
            return new Publisher(payload, () => {
                OfficeFileCommit.WriteAllBytes(path, payload, OfficeFileCommit.ConflictPolicy.FailIfExists);
                return new PublicationOutcome { Written = true, Path = path };
            });
        });
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PublishersHonorTheOtherAssemblyClaimAndRecoverAfterRelease(bool coreOwnsClaim) {
        Assert.NotSame(typeof(CoreCommitClaim).Assembly, typeof(ReaderCommitClaim).Assembly);
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-reader-foreign-claim-" + Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(folder, "publication.bin");
        byte[] payload = Encoding.UTF8.GetBytes("Published after the other assembly releases its claim.");
        Func<bool> publish = () => {
            if (coreOwnsClaim) return ReaderFileCommit.WriteAllBytes(destination, payload, overwrite: false);
            OfficeFileCommit.WriteAllBytes(destination, payload, OfficeFileCommit.ConflictPolicy.FailIfExists);
            return true;
        };

        try {
            Directory.CreateDirectory(folder);
            using (IDisposable claim = coreOwnsClaim
                       ? (IDisposable)CoreCommitClaim.Acquire(destination)
                       : ReaderCommitClaim.Acquire(destination)) {
                string liveClaim = Assert.Single(Directory.GetFileSystemEntries(folder));

                if (coreOwnsClaim) Assert.False(publish());
                else Assert.Throws<IOException>(() => publish());

                Assert.False(File.Exists(destination));
                Assert.Equal(new[] { liveClaim }, Directory.GetFileSystemEntries(folder));
            }

            Assert.Empty(Directory.GetFileSystemEntries(folder));
            Assert.True(publish());

            Assert.Equal(payload, File.ReadAllBytes(destination));
            Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(folder));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    private static Publisher CreateReaderPublisher(string materializer, string folder, int writer) {
        string identity = "writer-" + writer.ToString("D2", System.Globalization.CultureInfo.InvariantCulture);
        switch (materializer) {
            case "asset": {
                byte[] payload = Encoding.UTF8.GetBytes("Reader attachment from " + identity);
                var document = new OfficeDocumentReadResult {
                    Assets = new[] {
                        new OfficeDocumentAsset {
                            Id = identity,
                            Kind = "attachment",
                            FileName = "shared-export.bin",
                            Extension = ".bin",
                            PayloadBytes = payload
                        }
                    }
                };
                return new Publisher(payload, () => {
                    OfficeDocumentMaterializedAsset result = Assert.Single(document.WriteAssetsToDirectory(
                        folder,
                        new OfficeDocumentAssetMaterializationOptions { Overwrite = false, CreateDirectory = false }));
                    return new PublicationOutcome { Written = result.Written, Path = result.Path, SkippedReason = result.SkippedReason };
                });
            }
            case "table": {
                string csv = "Writer,Value\r\n" + identity + "," + writer.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var exports = new[] {
                    new ReaderTableExportBundle { Id = identity, FileNamePrefix = "shared-export", Csv = csv }
                };
                return new Publisher(Encoding.UTF8.GetBytes(csv), () => {
                    ReaderTableMaterializedExport result = Assert.Single(exports.WriteTableExportsToDirectory(
                        folder,
                        new ReaderTableExportMaterializationOptions {
                            Overwrite = false,
                            CreateDirectory = false,
                            IncludeCsv = true,
                            IncludeMarkdown = false,
                            IncludeJson = false
                        }));
                    return new PublicationOutcome { Written = result.Written, Path = result.Path, SkippedReason = result.SkippedReason };
                });
            }
            case "visual": {
                string mermaid = "graph TD\nW" + writer.ToString("D2", System.Globalization.CultureInfo.InvariantCulture) + "-->Destination";
                var exports = new[] {
                    new ReaderVisualExportBundle {
                        Id = identity,
                        FileNamePrefix = "shared-export",
                        PayloadExtension = ".mmd",
                        Payload = mermaid
                    }
                };
                return new Publisher(Encoding.UTF8.GetBytes(mermaid), () => {
                    ReaderVisualMaterializedExport result = Assert.Single(exports.WriteVisualExportsToDirectory(
                        folder,
                        new ReaderVisualExportMaterializationOptions {
                            Overwrite = false,
                            CreateDirectory = false,
                            IncludePayload = true,
                            IncludeJson = false
                        }));
                    return new PublicationOutcome { Written = result.Written, Path = result.Path, SkippedReason = result.SkippedReason };
                });
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(materializer));
        }
    }

    private static async Task AssertSinglePublisherAsync(string extension, Func<string, int, Publisher> createPublisher) {
        const int writerCount = 12;
        const int rounds = 4;
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-reader-publication-" + Guid.NewGuid().ToString("N"));
        try {
            for (int round = 0; round < rounds; round++) {
                string folder = Path.Combine(root, round.ToString(System.Globalization.CultureInfo.InvariantCulture));
                Directory.CreateDirectory(folder);
                string destination = Path.Combine(folder, "shared-export" + extension);
                Publisher[] publishers = Enumerable.Range(0, writerCount).Select(writer => createPublisher(folder, writer)).ToArray();
                using var ready = new CountdownEvent(writerCount);
                using var start = new ManualResetEventSlim(false);

                Task<PublicationOutcome>[] tasks = publishers.Select((publisher, writer) => Task.Factory.StartNew(() => {
                    ready.Signal();
                    start.Wait();
                    PublicationOutcome outcome;
                    try {
                        outcome = publisher.Publish();
                    } catch (IOException exception) {
                        outcome = new PublicationOutcome { Collision = exception };
                    }
                    outcome.Writer = writer;
                    return outcome;
                }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)).ToArray();

                // The start barrier occupies dedicated workers, leaving the constrained pool free.
                // Its timeout guards worker startup rather than measuring publication speed.
                bool allReady = ready.Wait(TimeSpan.FromSeconds(15));
                start.Set();
                PublicationOutcome[] outcomes = await Task.WhenAll(tasks);
                Assert.True(allReady, "All publication writers must reach the start barrier.");

                PublicationOutcome winner = Assert.Single(outcomes, outcome => outcome.Written);
                Assert.Equal(destination, winner.Path);
                Assert.Equal(publishers[winner.Writer].ExpectedBytes, File.ReadAllBytes(destination));
                Assert.All(outcomes.Where(outcome => !outcome.Written), outcome => {
                    if (outcome.Collision == null) {
                        Assert.False(string.IsNullOrWhiteSpace(outcome.SkippedReason));
                        Assert.Equal(destination, outcome.Path);
                    }
                });
                Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(folder));
                Directory.Delete(folder, recursive: true);
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    private sealed class Publisher {
        internal Publisher(byte[] expectedBytes, Func<PublicationOutcome> publish) {
            ExpectedBytes = expectedBytes;
            Publish = publish;
        }

        internal byte[] ExpectedBytes { get; }
        internal Func<PublicationOutcome> Publish { get; }
    }

    private sealed class PublicationOutcome {
        internal int Writer { get; set; }
        internal bool Written { get; set; }
        internal string? Path { get; set; }
        internal string? SkippedReason { get; set; }
        internal IOException? Collision { get; set; }
    }
}
