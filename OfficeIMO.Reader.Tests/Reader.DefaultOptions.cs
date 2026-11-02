using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderDefaultOptionsTests {
    [Theory]
    [InlineData("chunks-path")]
    [InlineData("chunks-stream")]
    [InlineData("chunks-bytes")]
    [InlineData("document-path")]
    [InlineData("document-stream")]
    [InlineData("document-bytes")]
    [InlineData("async-path")]
    [InlineData("async-stream")]
    [InlineData("async-bytes")]
    [InlineData("incremental-path")]
    [InlineData("incremental-stream")]
    [InlineData("ordered-batch")]
    [InlineData("detailed-batch")]
    [InlineData("completed-batch")]
#if NET8_0_OR_GREATER
    [InlineData("async-incremental-path")]
    [InlineData("async-incremental-stream")]
#endif
    public async Task DefaultInputAdmissionAppliesAcrossSourceRoutesAndExplicitPoliciesReplaceIt(string route) {
        string path = Path.Combine(Path.GetTempPath(), "reader-default-" + Guid.NewGuid().ToString("N") + ".txt");
        byte[] bytes = Encoding.UTF8.GetBytes(new string('a', 512));
        File.WriteAllBytes(path, bytes);
        try {
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithDefaultOptions(new ReaderOptions { MaxInputBytes = 128 }).Build();
            async Task Read(ReaderOptions? options) {
                using var stream = new MemoryStream(bytes);
                switch (route) {
                    case "chunks-path": reader.Read(path, options); break;
                    case "chunks-stream": reader.Read(stream, "source.txt", options); break;
                    case "chunks-bytes": reader.Read(bytes, "source.txt", options); break;
                    case "document-path": reader.ReadDocument(path, options); break;
                    case "document-stream": reader.ReadDocument(stream, "source.txt", options); break;
                    case "document-bytes": reader.ReadDocument(bytes, "source.txt", options); break;
                    case "async-path": await reader.ReadDocumentAsync(path, options); break;
                    case "async-stream": await reader.ReadDocumentAsync(stream, "source.txt", options); break;
                    case "async-bytes": await reader.ReadDocumentAsync(bytes, "source.txt", options); break;
                    case "incremental-path": reader.EnumerateChunks(path, options).ToArray(); break;
                    case "incremental-stream": reader.EnumerateChunks(stream, "source.txt", options).ToArray(); break;
                    case "ordered-batch": await reader.ReadDocumentsAsync(new[] { path }, options); break;
                    case "detailed-batch":
                        var detailed = await reader.ReadDocumentsDetailedAsync(new[] { path }, options);
                        var detailedError = detailed[0].Error;
                        if (detailedError != null) throw detailedError;
                        break;
                    case "completed-batch":
                        Exception? error = null;
                        await reader.ReadDocumentsAsCompletedAsync(new[] { path }, outcome => error = outcome.Error, options);
                        if (error != null) throw error;
                        break;
#if NET8_0_OR_GREATER
                    case "async-incremental-path": await foreach (var _ in reader.EnumerateChunksAsync(path, options)) { } break;
                    case "async-incremental-stream": await foreach (var _ in reader.EnumerateChunksAsync(stream, "source.txt", options)) { } break;
#endif
                }
            }
            await Assert.ThrowsAnyAsync<IOException>(() => Read(null));
            await Read(new ReaderOptions { MaxInputBytes = 1024 });
            Assert.Equal(128, reader.DefaultOptions.MaxInputBytes);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void RegistrationBuildAndReturnedProfilesHaveIndependentNestedSnapshots() {
        var source = new ReaderOptions { ComputeHashes = false, ResourceLimits = new ReaderResourceLimits { MaxChunks = 2 } };
        var builder = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithDefaultOptions(source);
        source.ComputeHashes = true; source.ResourceLimits.MaxChunks = 0;
        var first = builder.Build();
        var returned = first.DefaultOptions;
        returned.ComputeHashes = true; returned.ResourceLimits!.MaxChunks = 0;
        Assert.False(first.DefaultOptions.ComputeHashes);
        Assert.Equal(2, first.DefaultOptions.ResourceLimits!.MaxChunks);
        Assert.Null(first.ReadDocument(Encoding.UTF8.GetBytes("evidence"), "source.txt").Source.SourceHash);
        builder.WithDefaultOptions(new ReaderOptions { MaxInputBytes = 1 });
        var second = builder.Build();
        Assert.ThrowsAny<IOException>(() => second.ReadDocument(Encoding.UTF8.GetBytes("evidence"), "source.txt"));
        Assert.Null(first.ReadDocument(Encoding.UTF8.GetBytes("evidence"), "source.txt").Source.SourceHash);
    }

    [Fact]
    public void FolderDefaultsApplyToEnumerationAndIngestionAndExplicitFolderPoliciesReplaceThem() {
        string directory = Path.Combine(Path.GetTempPath(), "reader-default-folder-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(directory, "child"));
        try {
            File.WriteAllText(Path.Combine(directory, "first.txt"), "first");
            File.WriteAllText(Path.Combine(directory, "child", "second.txt"), "second");
            var source = new ReaderOptions { FolderOptions = new ReaderFolderOptions { Recurse = false, MaxFiles = 1 } };
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithDefaultOptions(source).Build();
            source.FolderOptions.Recurse = true;
            reader.DefaultOptions.FolderOptions!.Recurse = true;
            Assert.Single(reader.EnumerateDocumentPaths(new[] { directory }));
            Assert.Equal(1, reader.ReadFolderDetailed(directory).FilesParsed);
            Assert.Equal(2, reader.ReadFolderDetailed(directory, new ReaderFolderOptions { Recurse = true, MaxFiles = 10 }).FilesParsed);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void CapturingDefaultsInsideALeanProcessorDoesNotChangeTheConfiguredProjection() {
        OfficeDocumentReader? captured = null;
        var parent = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddProcessor(new DelegateOfficeDocumentProcessor("configure", (document, _) => {
            captured = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithDefaultOptions(new ReaderOptions { Projection = ReaderProjection.Full }).Build();
            return document;
        })).Build();
        parent.ReadDocument(Encoding.UTF8.GetBytes("evidence"), "source.txt", new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        Assert.NotNull(captured);
        Assert.Equal(ReaderProjection.Full, captured!.DefaultOptions.Projection);
    }

    [Fact]
    public void DefaultsDoNotChangeStandaloneExtensionDetectionIntoAFileRead() {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithDefaultOptions(new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent }).Build();
        Assert.Equal(ReaderInputKind.Word, reader.DetectKind("nonexistent.docx"));
    }
}
