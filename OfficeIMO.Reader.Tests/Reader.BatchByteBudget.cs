using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Reader.IWork;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderBatchByteBudgetTests {
    [Theory]
    [InlineData("ordered")]
    [InlineData("detailed")]
    [InlineData("completed")]
    public async Task AggregateAdmissionUsesRemainingBytesAndKeepsResilientOutcomes(string route) {
        string directory = NewDirectory();
        try {
            string[] paths = { Path.Combine(directory, "large.txt"), Path.Combine(directory, "two.txt"), Path.Combine(directory, "three.txt") };
            File.WriteAllBytes(paths[0], new byte[100]); File.WriteAllBytes(paths[1], new byte[2]); File.WriteAllBytes(paths[2], new byte[3]);
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithMaxConcurrentReads(4).Build();
            var batch = new ReaderBatchOptions { MaxTotalBytes = 5, MaxDegreeOfParallelism = 4 };
            if (route == "ordered") {
                await Assert.ThrowsAsync<IOException>(() => reader.ReadDocumentsAsync(paths, batchOptions: batch));
            } else {
                var outcomes = new List<ReaderDocumentReadOutcome>();
                if (route == "detailed") outcomes.AddRange(await reader.ReadDocumentsDetailedAsync(paths, batchOptions: batch));
                else await reader.ReadDocumentsAsCompletedAsync(paths, outcomes.Add, batchOptions: batch);
                Assert.Equal(paths, outcomes.Select(outcome => outcome.Path));
                Assert.IsAssignableFrom<IOException>(outcomes[0].Error);
                Assert.Equal(5, outcomes.Where(outcome => outcome.Succeeded).Sum(outcome => outcome.Document!.Source.LengthBytes));
                Assert.Equal(2, outcomes.Count(outcome => outcome.Succeeded));
            }
            Assert.Equal(4, batch.MaxDegreeOfParallelism);
            Assert.Equal(5, batch.MaxTotalBytes);
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("ordered")]
    [InlineData("detailed")]
    [InlineData("completed")]
    public async Task WholeReadResourceBudgetsStillSpanByteBoundedSources(string route) {
        string directory = NewDirectory();
        try {
            string[] paths = { Path.Combine(directory, "first.txt"), Path.Combine(directory, "second.txt") };
            foreach (string path in paths) File.WriteAllText(path, "evidence");
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
            var options = new ReaderOptions { ResourceLimits = new ReaderResourceLimits { MaxChunks = 1 } };
            var batch = new ReaderBatchOptions { MaxTotalBytes = 100 };
            await Assert.ThrowsAsync<ReaderResourceLimitException>(async () => {
                if (route == "ordered") await reader.ReadDocumentsAsync(paths, options, batch);
                else if (route == "detailed") await reader.ReadDocumentsDetailedAsync(paths, options, batch);
                else await reader.ReadDocumentsAsCompletedAsync(paths, _ => { }, options, batch);
            });
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task AParentAllowanceDoesNotRelaxARegisteredNativeDefault() {
        string directory = NewDirectory();
        try {
            string path = Path.Combine(directory, "source.bounded"); File.WriteAllBytes(path, new byte[32]);
            bool entered = false;
            var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
                Id = "bounded", Extensions = new[] { ".bounded" }, DefaultMaxInputBytes = 16,
                ReadDocumentPath = (_, _, _) => { entered = true; return new OfficeDocumentReadResult(); }
            }).Build();
            var outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }, batchOptions: new ReaderBatchOptions { MaxTotalBytes = 100 }));
            Assert.False(entered); Assert.IsAssignableFrom<IOException>(outcome.Error);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task DirectoryPackagesUseTheirNativeSnapshotLengthWithinTheSameBatch() {
        string directory = NewDirectory();
        try {
            string bundle = Path.Combine(directory, "evidence.pages");
            ZipFile.ExtractToDirectory(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "nim-iwork", "simple.pages"), bundle);
            var reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().AddPlainTextHandlers().Build();
            long bytes = reader.ReadDocument(bundle).Source.LengthBytes!.Value;
            string following = Path.Combine(directory, "following.txt"); File.WriteAllText(following, "too much");
            var outcomes = await reader.ReadDocumentsDetailedAsync(new[] { bundle, following }, batchOptions: new ReaderBatchOptions { MaxTotalBytes = bytes });
            Assert.True(outcomes[0].Succeeded);
            Assert.Equal(bytes, outcomes[0].Document!.Source.LengthBytes);
            Assert.Contains("hello pages", string.Join("\n", outcomes[0].Document!.Chunks.Select(chunk => chunk.Text)));
            Assert.False(outcomes[1].Succeeded);
            Assert.Contains("MaxTotalBytes", outcomes[1].Error!.Message);
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(0, 2)]
    [InlineData(100, 0)]
    public async Task InvalidBatchLimitsRemainErrorsWhenByteAccountingIsEnabled(long maximum, int degree) {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => reader.ReadDocumentsAsync(Array.Empty<string>(),
            batchOptions: new ReaderBatchOptions { MaxTotalBytes = maximum, MaxDegreeOfParallelism = degree }));
    }

    private static string NewDirectory() {
        string path = Path.Combine(Path.GetTempPath(), "reader-batch-bytes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path); return path;
    }
}
