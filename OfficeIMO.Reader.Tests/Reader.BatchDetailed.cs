using OfficeIMO.Reader;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderBatchDetailedTests {
    [Fact]
    public async Task DetailedBatchCapturesPerPathFailuresAndStreamsCompletedOutcomes() {
        const string extension = ".readerbatchdetail";
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-batch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string[] paths = {
            Path.Combine(root, "01-good" + extension),
            Path.Combine(root, "02-bad" + extension),
            Path.Combine(root, "03-good" + extension)
        };
        foreach (string path in paths) File.WriteAllText(path, "input");

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .WithMaxConcurrentReads(2)
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.batch.detailed",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPathAsync = (path, options, cancellationToken) => {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (Path.GetFileName(path).Contains("bad", StringComparison.Ordinal)) {
                            throw new InvalidDataException("Synthetic per-path failure.");
                        }

                        return Task.FromResult(CreateResult(path));
                    }
                })
                .Build();
            var completed = new ConcurrentBag<ReaderDocumentReadOutcome>();

            IReadOnlyList<ReaderDocumentReadOutcome> outcomes = await reader.ReadDocumentsDetailedAsync(
                paths,
                batchOptions: new ReaderBatchOptions {
                    MaxDocuments = 3,
                    MaxDegreeOfParallelism = 2
                },
                onCompleted: completed.Add);

            Assert.Equal(paths, outcomes.Select(outcome => outcome.Path).ToArray());
            Assert.Equal(new[] { 0, 1, 2 }, outcomes.Select(outcome => outcome.Index).ToArray());
            Assert.Equal(2, outcomes.Count(outcome => outcome.Succeeded));
            ReaderDocumentReadOutcome failed = Assert.Single(outcomes, outcome => !outcome.Succeeded);
            Assert.IsType<InvalidDataException>(failed.Error);
            Assert.Null(failed.Document);
            Assert.Equal(3, completed.Count);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void PathEnumerationUsesRegisteredFormatsAndCallerControlledFolderLimits() {
        const string extension = ".readerpathexpand";
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-paths-" + Guid.NewGuid().ToString("N"));
        string nested = Path.Combine(root, "nested");
        Directory.CreateDirectory(nested);
        string first = Path.Combine(root, "01-first" + extension);
        string second = Path.Combine(nested, "02-second" + extension);
        string ignored = Path.Combine(root, "ignored.unsupported");
        File.WriteAllText(first, "first");
        File.WriteAllText(second, "second");
        File.WriteAllText(ignored, "ignored");

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.path.enumeration",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPath = (path, options, cancellationToken) => CreateResult(path)
                })
                .Build();

            string[] topLevel = reader.EnumerateDocumentPaths(
                new[] { root },
                new ReaderFolderOptions { Recurse = false, MaxFiles = 10 }).ToArray();
            string[] recursive = reader.EnumerateDocumentPaths(
                new[] { root },
                new ReaderFolderOptions { Recurse = true, MaxFiles = int.MaxValue }).ToArray();
            string[] byteBounded = reader.EnumerateDocumentPaths(
                new[] { root },
                new ReaderFolderOptions {
                    Recurse = true,
                    MaxFiles = int.MaxValue,
                    MaxTotalBytes = new FileInfo(first).Length
                }).ToArray();
            string[] explicitUnsupported = reader.EnumerateDocumentPaths(new[] { ignored }).ToArray();

            Assert.Equal(new[] { first }, topLevel);
            Assert.Equal(new[] { first, second }, recursive);
            Assert.Equal(new[] { first }, byteBounded);
            Assert.Equal(new[] { ignored }, explicitUnsupported);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task AsCompletedBatchReportsOutcomesWithoutReturningAMaterializedCollection() {
        const string extension = ".readerbatchstream";
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-stream-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string[] paths = Enumerable.Range(1, 8)
            .Select(index => Path.Combine(root, index.ToString("00") + extension))
            .ToArray();
        foreach (string path in paths) File.WriteAllText(path, "input");

        try {
            int activeReads = 0;
            int peakReads = 0;
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .WithMaxConcurrentReads(6)
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.batch.as-completed",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPathAsync = async (path, options, cancellationToken) => {
                        int active = Interlocked.Increment(ref activeReads);
                        UpdateMaximum(ref peakReads, active);
                        try {
                            await Task.Delay(50, cancellationToken).ConfigureAwait(false);
                            return CreateResult(path);
                        } finally {
                            Interlocked.Decrement(ref activeReads);
                        }
                    }
                })
                .Build();
            var completed = new ConcurrentBag<ReaderDocumentReadOutcome>();

            await reader.ReadDocumentsAsCompletedAsync(
                paths,
                completed.Add,
                batchOptions: new ReaderBatchOptions {
                    MaxDocuments = paths.Length,
                    MaxDegreeOfParallelism = 6
                });

            Assert.Equal(paths.Length, completed.Count);
            Assert.All(completed, outcome => Assert.True(outcome.Succeeded));
            Assert.Equal(paths, completed.OrderBy(outcome => outcome.Index).Select(outcome => outcome.Path).ToArray());
            Assert.True(peakReads >= 5, $"Expected at least five concurrent reads but observed {peakReads}.");
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void UpdateMaximum(ref int target, int value) {
        int current;
        do {
            current = Volatile.Read(ref target);
            if (value <= current) return;
        } while (Interlocked.CompareExchange(ref target, value, current) != current);
    }

    private static OfficeDocumentReadResult CreateResult(string path) => new OfficeDocumentReadResult {
        Kind = ReaderInputKind.Text,
        Source = new OfficeDocumentSource { Path = path },
        Blocks = new[] {
            new OfficeDocumentBlock {
                Id = "block:0001",
                Kind = "paragraph",
                Text = "searchable text",
                Location = new ReaderLocation { Path = path }
            }
        }
    };
}
